<?php

use FluentMail\App\Services\Mailer\Providers\Outlook\API as OutlookAPI;
use FluentMail\App\Services\Mailer\Providers\Outlook\Handler as OutlookHandler;
use FluentMail\Includes\Support\ValidationException;

return function () {
    $validSettings = function ($sender = 'sender@example.test') {
        return [
            'provider'      => 'outlook',
            'auth_mode'     => OutlookHandler::AUTH_APP_ONLY,
            'sender_email'  => $sender,
            'key_store'     => 'db',
            'tenant_id'     => '11111111-1111-4111-8111-111111111111',
            'client_id'     => '22222222-2222-4222-8222-222222222222',
            'client_secret' => 'obviously-fake-suite-secret',
            'access_token'  => 'obviously-fake-cached-token',
            'expire_stamp'  => time() + 900,
        ];
    };

    $withWriteFuses = function (callable $callback) {
        $optionFuse = function ($newValue, $oldValue) {
            return $oldValue;
        };
        $scheduleFuse = function () {
            return false;
        };

        add_filter('pre_update_option_fluentmail-settings', $optionFuse, PHP_INT_MAX, 2);
        add_filter('pre_schedule_event', $scheduleFuse, PHP_INT_MAX, 3);
        try {
            return $callback();
        } finally {
            remove_filter('pre_update_option_fluentmail-settings', $optionFuse, PHP_INT_MAX);
            remove_filter('pre_schedule_event', $scheduleFuse, PHP_INT_MAX);
            wp_cache_delete('fluentmail-settings', 'options');
            wp_cache_delete('alloptions', 'options');
        }
    };

    FsmtpTest::case('Outlook connections without an authentication mode remain delegated', function () {
        FsmtpTest::assertSame(
            OutlookHandler::AUTH_DELEGATED,
            OutlookHandler::getAuthMode(['provider' => 'outlook']),
            'legacy Outlook authentication mode'
        );
    });

    FsmtpTest::case('Outlook mode transitions discard incompatible token state', function () use ($validSettings) {
        $handler = new OutlookHandler();
        $delegated = [
            'provider'      => 'outlook',
            'sender_email'  => 'sender@example.test',
            'key_store'     => 'db',
            'client_id'     => '22222222-2222-4222-8222-222222222222',
            'client_secret' => 'obviously-fake-suite-secret',
            'auth_token'    => 'obviously-fake-auth-code',
            'access_token'  => 'obviously-fake-delegated-token',
            'refresh_token' => 'obviously-fake-refresh-token',
            'expire_stamp'  => time() + 900,
        ];

        $appOnly = $handler->prepareConnectionForSave($validSettings(), $delegated);
        FsmtpTest::assert(!isset($appOnly['auth_token']), 'app-only auth code was retained');
        FsmtpTest::assert(!isset($appOnly['refresh_token']), 'app-only refresh token was retained');
        FsmtpTest::assert(!isset($appOnly['access_token']), 'delegated access token crossed into app-only mode');

        $delegated['auth_mode'] = OutlookHandler::AUTH_DELEGATED;
        $backToDelegated = $handler->prepareConnectionForSave($delegated, $validSettings());
        FsmtpTest::assert(!isset($backToDelegated['tenant_id']), 'tenant ID crossed into delegated mode');
        FsmtpTest::assert(!isset($backToDelegated['access_token']), 'app-only access token crossed into delegated mode');
    });

    FsmtpTest::case('App-only token request uses tenant endpoint and exact client credentials parameters', function () {
        $tenant = '11111111-1111-4111-8111-111111111111';
        $client = '22222222-2222-4222-8222-222222222222';
        $secret = 'obviously-fake-suite-secret';

        FsmtpTest::interceptHttp(function ($url) use ($tenant) {
            if ($url === 'https://login.microsoftonline.com/' . $tenant . '/oauth2/v2.0/token') {
                return [
                    'headers' => [],
                    'body' => wp_json_encode([
                        'token_type' => 'Bearer',
                        'expires_in' => 3600,
                        'access_token' => 'obviously-fake-new-token',
                    ]),
                    'response' => ['code' => 200, 'message' => 'OK'],
                    'cookies' => [],
                    'filename' => null,
                ];
            }
            return null;
        });

        $tokens = (new OutlookAPI($client, $secret))->requestAppToken($tenant);
        FsmtpTest::assert(!is_wp_error($tokens), 'client credentials request returned an error');
        FsmtpTest::assertSame(
            ['access_token' => 'obviously-fake-new-token', 'expires_in' => 3600],
            $tokens,
            'normalized app-only token response'
        );

        $requests = FsmtpTest::httpRequests();
        FsmtpTest::assertSame(1, count($requests), 'client credentials request count');
        parse_str(isset($requests[0]['args']['body']) ? $requests[0]['args']['body'] : '', $body);
        FsmtpTest::assertSame('client_credentials', isset($body['grant_type']) ? $body['grant_type'] : null, 'grant type');
        FsmtpTest::assertSame('https://graph.microsoft.com/.default', isset($body['scope']) ? $body['scope'] : null, 'Graph scope');
        FsmtpTest::assertSame($client, isset($body['client_id']) ? $body['client_id'] : null, 'client ID');
        FsmtpTest::assertSame($secret, isset($body['client_secret']) ? $body['client_secret'] : null, 'client secret value');
        FsmtpTest::assert(!isset($body['redirect_uri']), 'client credentials request included a redirect URI');
        FsmtpTest::assert(!isset($tokens['refresh_token']), 'client credentials response exposed a refresh token');
    });

    FsmtpTest::case('App-only token errors are useful and redact credentials and tokens', function () {
        $secret = 'obviously-fake-suite-secret';
        $leakedToken = 'eyJ.obviously-fake-token-material-for-redaction.signature';
        FsmtpTest::interceptHttp(function () use ($secret, $leakedToken) {
            return [
                'headers' => [],
                'body' => wp_json_encode([
                    'error' => 'invalid_client',
                    'error_description' => 'Credential ' . $secret . ' failed with Bearer ' . $leakedToken,
                ]),
                'response' => ['code' => 401, 'message' => 'Unauthorized'],
                'cookies' => [],
                'filename' => null,
            ];
        });

        $result = (new OutlookAPI('22222222-2222-4222-8222-222222222222', $secret))
            ->requestAppToken('11111111-1111-4111-8111-111111111111');

        FsmtpTest::assert(is_wp_error($result), 'rejected app credentials did not return WP_Error');
        $message = is_wp_error($result) ? $result->get_error_message() : '';
        FsmtpTest::assert(strpos($message, 'Credential') !== false, 'sanitized token error lost useful context');
        FsmtpTest::assert(strpos($message, $secret) === false, 'token error leaked client secret');
        FsmtpTest::assert(strpos($message, $leakedToken) === false, 'token error leaked bearer token');
    });

    FsmtpTest::case('Delegated and app-only Graph sends preserve MIME and select distinct endpoints', function () {
        $mime = base64_encode(
            "From: Sender <sender@example.test>\r\n" .
            "To: Recipient <recipient@example.test>\r\n" .
            "Cc: Copy <copy@example.test>\r\n" .
            "Bcc: Blind <blind@example.test>\r\n" .
            "Reply-To: Replies <replies@example.test>\r\n" .
            "Content-Type: multipart/mixed; boundary=suite\r\n" .
            "X-Suite-Header: retained\r\n\r\n" .
            "--suite\r\nContent-Type: multipart/alternative\r\n\r\n" .
            "plain body\r\n<html><body>html body</body></html>\r\n" .
            "--suite\r\nContent-Disposition: attachment; filename=fixture.txt\r\n\r\nattachment body\r\n--suite--"
        );

        FsmtpTest::interceptHttp(function () {
            return [
                'headers' => ['request-id' => 'obviously-fake-request-id'],
                'body' => '',
                'response' => ['code' => 202, 'message' => 'Accepted'],
                'cookies' => [],
                'filename' => null,
            ];
        });

        $api = new OutlookAPI('22222222-2222-4222-8222-222222222222', 'obviously-fake-suite-secret');
        $api->sendMime($mime, 'obviously-fake-delegated-token');
        $api->sendMime($mime, 'obviously-fake-app-token', 'sender+route@example.test');

        $requests = FsmtpTest::httpRequests();
        FsmtpTest::assertSame('https://graph.microsoft.com/v1.0/me/sendMail', $requests[0]['url'], 'delegated Graph endpoint');
        FsmtpTest::assertSame(
            'https://graph.microsoft.com/v1.0/users/sender%2Broute%40example.test/sendMail',
            $requests[1]['url'],
            'app-only Graph endpoint'
        );
        FsmtpTest::assertSame($mime, $requests[0]['args']['body'], 'delegated MIME body');
        FsmtpTest::assertSame($mime, $requests[1]['args']['body'], 'app-only MIME body');
        FsmtpTest::assertSame('text/plain', $requests[1]['args']['headers']['Content-Type'], 'MIME content type');
    });

    FsmtpTest::case('Graph errors are sanitized without hiding Microsoft diagnostics', function () {
        $secret = 'obviously-fake-suite-secret';
        $token = 'obviously-fake-app-token';
        FsmtpTest::interceptHttp(function () use ($secret, $token) {
            return [
                'headers' => [],
                'body' => wp_json_encode([
                    'error' => [
                        'code' => 'ErrorAccessDenied',
                        'message' => 'Mailbox authorization denied; client_secret=' . $secret . ' Bearer ' . $token,
                    ],
                ]),
                'response' => ['code' => 403, 'message' => 'Forbidden'],
                'cookies' => [],
                'filename' => null,
            ];
        });

        $result = (new OutlookAPI('22222222-2222-4222-8222-222222222222', $secret))
            ->sendMime('obviously-fake-mime', $token, 'sender@example.test');
        FsmtpTest::assert(is_wp_error($result), 'denied Graph send did not return WP_Error');
        $message = is_wp_error($result) ? $result->get_error_message() : '';
        FsmtpTest::assert(strpos($message, 'Mailbox authorization denied') !== false, 'Graph diagnostic was lost');
        FsmtpTest::assert(strpos($message, $secret) === false, 'Graph error leaked client secret');
        FsmtpTest::assert(strpos($message, $token) === false, 'Graph error leaked access token');
    });

    FsmtpTest::case('App-only cached tokens are reused and expired tokens are renewed once', function () use (
        $validSettings,
        $withWriteFuses
    ) {
        FsmtpTest::interceptHttp(function ($url) {
            if (strpos($url, 'login.microsoftonline.com/11111111-1111-4111-8111-111111111111/oauth2/v2.0/token') !== false) {
                return [
                    'headers' => [],
                    'body' => wp_json_encode([
                        'access_token' => 'obviously-fake-renewed-token',
                        'expires_in' => 3600,
                    ]),
                    'response' => ['code' => 200, 'message' => 'OK'],
                    'cookies' => [],
                    'filename' => null,
                ];
            }
            return null;
        });

        $method = new ReflectionMethod(OutlookHandler::class, 'getAccessToken');
        $method->setAccessible(true);
        $handler = new OutlookHandler();
        $future = $validSettings('future@example.test');

        FsmtpTest::assertSame(
            'obviously-fake-cached-token',
            $method->invoke($handler, $future, false),
            'cached app-only token'
        );
        FsmtpTest::assertSame(0, count(FsmtpTest::httpRequests()), 'future app-only token request count');

        $expired = array_merge($future, [
            'sender_email' => 'expired@example.test',
            'expire_stamp' => time() - 1,
        ]);
        $handler->setSettings($expired);
        $renewed = $withWriteFuses(function () use ($method, $handler, $expired) {
            return $method->invoke($handler, $expired, false);
        });

        FsmtpTest::assertSame('obviously-fake-renewed-token', $renewed, 'renewed app-only token');
        FsmtpTest::assertSame(1, count(FsmtpTest::httpRequests()), 'expired app-only token request count');
        parse_str(FsmtpTest::httpRequests()[0]['args']['body'], $body);
        FsmtpTest::assert(!isset($body['refresh_token']), 'app-only renewal included a refresh token');

        $second = $method->invoke($handler, $handler->getSetting(), false);
        FsmtpTest::assertSame('obviously-fake-renewed-token', $second, 'batch-reused app-only token');
        FsmtpTest::assertSame(1, count(FsmtpTest::httpRequests()), 'batch app-only token request count');
    });

    FsmtpTest::case('App-only validates tenant ID client ID and organizational sender', function () use ($validSettings) {
        $invalid = array_merge($validSettings(), [
            'tenant_id' => 'common',
            'client_id' => 'not-a-client-id',
            'sender_email' => 'not-a-mailbox',
        ]);

        try {
            (new OutlookHandler())->validateProviderInformation($invalid);
            FsmtpTest::fail('invalid app-only identifiers passed validation');
        } catch (ValidationException $e) {
            $errors = $e->errors();
            FsmtpTest::assert(isset($errors['tenant_id']), 'tenant ID validation error missing');
            FsmtpTest::assert(isset($errors['client_id']), 'client ID validation error missing');
            FsmtpTest::assert(isset($errors['sender_email']), 'sender validation error missing');
        }
    });

    FsmtpTest::case('Outlook access and refresh tokens are encrypted on new settings writes', function () use ($validSettings) {
        $original = get_option('fluentmail-settings', null);
        $sender = 'encrypted-' . FsmtpTest::uniq() . '@example.test';
        $connectionKey = md5($sender);
        $settings = [
            'connections' => [
                $connectionKey => [
                    'title' => 'Suite encrypted Outlook',
                    'provider_settings' => array_merge($validSettings($sender), [
                        'access_token' => 'obviously-fake-access-for-encryption',
                        'refresh_token' => 'obviously-fake-refresh-for-legacy-coverage',
                        'disable_encryption' => 'yes',
                    ]),
                ],
            ],
            'mappings' => [$sender => $connectionKey],
            'misc' => ['default_connection' => $connectionKey],
        ];

        try {
            fluentMailSetSettings($settings);
            $raw = get_option('fluentmail-settings');
            $rawProvider = $raw['connections'][$connectionKey]['provider_settings'];
            FsmtpTest::assertSame('yes', $rawProvider['outlook_tokens_encrypted'], 'Outlook token encryption marker');
            FsmtpTest::assertSame(
                'obviously-fake-suite-secret',
                $rawProvider['client_secret'],
                'client-secret encryption opt-out'
            );
            FsmtpTest::assert(
                $rawProvider['access_token'] !== 'obviously-fake-access-for-encryption',
                'Outlook access token remained plaintext'
            );
            FsmtpTest::assert(
                $rawProvider['refresh_token'] !== 'obviously-fake-refresh-for-legacy-coverage',
                'Outlook refresh token remained plaintext'
            );

            $decoded = fluentMailGetSettings([], false);
            FsmtpTest::assertSame(
                'obviously-fake-access-for-encryption',
                $decoded['connections'][$connectionKey]['provider_settings']['access_token'],
                'decrypted Outlook access token'
            );
            FsmtpTest::assertSame(
                'obviously-fake-refresh-for-legacy-coverage',
                $decoded['connections'][$connectionKey]['provider_settings']['refresh_token'],
                'decrypted Outlook refresh token'
            );
        } finally {
            if ($original === null) {
                delete_option('fluentmail-settings');
            } else {
                update_option('fluentmail-settings', $original);
            }
            wp_cache_delete('fluentmail-settings', 'options');
            wp_cache_delete('alloptions', 'options');
            fluentMailGetSettings([], false);
        }
    });

    FsmtpTest::case('Legacy Outlook plaintext tokens remain readable before their next settings write', function () {
        $original = get_option('fluentmail-settings', null);
        $sender = 'legacy-' . FsmtpTest::uniq() . '@example.test';
        $connectionKey = md5($sender);
        $legacy = [
            'use_encrypt' => 'yes',
            'test' => fluentMailEncryptDecrypt('test', 'e'),
            'connections' => [
                $connectionKey => [
                    'title' => 'Suite legacy Outlook',
                    'provider_settings' => [
                        'provider' => 'outlook',
                        'sender_email' => $sender,
                        'key_store' => 'db',
                        'client_secret' => fluentMailEncryptDecrypt('obviously-fake-legacy-secret', 'e'),
                        'access_token' => 'obviously-fake-legacy-access',
                        'refresh_token' => 'obviously-fake-legacy-refresh',
                    ],
                ],
            ],
            'mappings' => [$sender => $connectionKey],
            'misc' => ['default_connection' => $connectionKey],
        ];

        try {
            update_option('fluentmail-settings', $legacy);
            wp_cache_delete('fluentmail-settings', 'options');
            wp_cache_delete('alloptions', 'options');
            $decoded = fluentMailGetSettings([], false);
            $provider = $decoded['connections'][$connectionKey]['provider_settings'];

            FsmtpTest::assertSame('obviously-fake-legacy-secret', $provider['client_secret'], 'legacy client secret');
            FsmtpTest::assertSame('obviously-fake-legacy-access', $provider['access_token'], 'legacy access token');
            FsmtpTest::assertSame('obviously-fake-legacy-refresh', $provider['refresh_token'], 'legacy refresh token');
            FsmtpTest::assertSame(
                OutlookHandler::AUTH_DELEGATED,
                OutlookHandler::getAuthMode($provider),
                'legacy delegated mode'
            );
        } finally {
            if ($original === null) {
                delete_option('fluentmail-settings');
            } else {
                update_option('fluentmail-settings', $original);
            }
            wp_cache_delete('fluentmail-settings', 'options');
            wp_cache_delete('alloptions', 'options');
            fluentMailGetSettings([], false);
        }
    });
};
