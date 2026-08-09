<?php

namespace FluentMail\App\Services\Mailer\Providers\Outlook;

use FluentMail\Includes\Support\Arr;

class API
{
    private $clientId;
    private $clientSecret;

    public function __construct($clientId = '', $clientSecret = '')
    {
        $this->clientId     = $clientId;
        $this->clientSecret = $clientSecret;
    }

    public function getAuthUrl()
    {

        $fluentClient = new \FluentMail\Includes\OAuth2Provider($this->getConfig());

        return $fluentClient->getAuthorizationUrl();

    }

    public function generateToken($authCode)
    {
        return $this->sendTokenRequest('authorization_code', [
            'code' => $authCode
        ]);
    }

    /**
     * @return mixed|string
     */
    public function sendTokenRequest($type, $params)
    {
        $fluentClient = new \FluentMail\Includes\OAuth2Provider($this->getConfig());
        try {
            $tokens = $fluentClient->getAccessToken($type, $params);
            return $tokens;
        } catch (\Exception$exception) {
            return new \WP_Error(422, $exception->getMessage());
        }
    }

    /**
     * @return array | \WP_Error
     */
    public function requestAppToken($tenantId)
    {
        $url = 'https://login.microsoftonline.com/' . rawurlencode($tenantId) . '/oauth2/v2.0/token';
        $body = http_build_query([
            'grant_type'    => 'client_credentials',
            'scope'         => 'https://graph.microsoft.com/.default',
            'client_id'     => $this->clientId,
            'client_secret' => $this->clientSecret,
        ], '', '&', PHP_QUERY_RFC3986);

        $response = wp_remote_request($url, [
            'method'  => 'POST',
            'headers' => [
                'content-type' => 'application/x-www-form-urlencoded'
            ],
            'body'    => $body
        ]);

        if (is_wp_error($response)) {
            return new \WP_Error(422, $this->sanitizeError($response->get_error_message()));
        }

        $responseBody = json_decode(wp_remote_retrieve_body($response), true);
        $responseCode = wp_remote_retrieve_response_code($response);

        if ($responseCode >= 300 || !is_array($responseBody) || empty($responseBody['access_token'])) {
            $message = Arr::get($responseBody, 'error_description');
            if (!$message) {
                $message = Arr::get($responseBody, 'error');
            }
            if (!$message) {
                $message = __('Microsoft did not return a usable application access token.', 'fluent-smtp');
            }

            return new \WP_Error($responseCode ?: 422, $this->sanitizeError($message));
        }

        return [
            'access_token' => $responseBody['access_token'],
            'expires_in'   => !empty($responseBody['expires_in']) ? (int)$responseBody['expires_in'] : 3600,
        ];
    }

    public function sendMime($mime, $accessToken, $sender = '')
    {
        $endpoint = 'https://graph.microsoft.com/v1.0/me/sendMail';
        if ($sender) {
            $endpoint = 'https://graph.microsoft.com/v1.0/users/' . rawurlencode($sender) . '/sendMail';
        }

        $response = wp_remote_request($endpoint, [
            'method'  => 'POST',
            'headers' => [
                'Authorization' => 'Bearer ' . $accessToken,
                'Content-Type'  => 'text/plain'
            ],
            'body'    => $mime
        ]);

        if (is_wp_error($response)) {
            return new \WP_Error(
                $response->get_error_code(),
                $this->sanitizeError($response->get_error_message(), $accessToken)
            );
        }

        $responseCode = wp_remote_retrieve_response_code($response);

        if ($responseCode >= 300) {
            /*
             * Graph's own message first. The other candidate is the HTTP
             * reason phrase, which is set on essentially every failure and so
             * always won the old ordering - that is where the useless
             * "Unauthorized" in the logs came from, while the body sitting
             * right next to it said the access token had expired and the
             * account needed reconnecting.
             */
            $responseBody = json_decode(wp_remote_retrieve_body($response), true);

            $error = Arr::get($responseBody, 'error.message');

            if (!$error) {
                $error = Arr::get($response, 'response.message');
            }

            if (!$error) {
                $error = __('Something went wrong with the Outlook API. Please check your API Settings', 'fluent-smtp');
            }

            return new \WP_Error($responseCode, $this->sanitizeError($error, $accessToken));
        }

        $header = wp_remote_retrieve_headers($response);

        return is_object($header) && method_exists($header, 'getAll')
            ? $header->getAll()
            : (array)$header;
    }

    /**
     * Keep Microsoft diagnostics useful without reflecting credentials or
     * bearer tokens into FluentSMTP logs and notices.
     *
     * @param mixed $message
     * @param string $accessToken
     * @return string
     */
    private function sanitizeError($message, $accessToken = '')
    {
        $message = wp_strip_all_tags((string)$message);

        foreach (array_filter([$this->clientSecret, $accessToken]) as $credential) {
            $message = str_replace($credential, '[redacted]', $message);
        }

        $message = preg_replace('/(Bearer\s+)[A-Za-z0-9._~+\/-]+/i', '$1[redacted]', $message);
        $message = preg_replace('/((?:client_secret|access_token|refresh_token)=)[^&\s]+/i', '$1[redacted]', $message);
        $message = preg_replace('/\beyJ[A-Za-z0-9._-]{20,}\b/', '[redacted token]', $message);

        return trim((string)$message);
    }

    public function getRedirectUrl()
    {
        return rest_url('fluent-smtp/outlook_callback');
    }

    private function getConfig()
    {
        return [
            'clientId'                => $this->clientId,
            'clientSecret'            => $this->clientSecret,
            'redirectUri'             => $this->getRedirectUrl(),
            'urlAuthorize'            => 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
            'urlAccessToken'          => 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
            'urlResourceOwnerDetails' => '',
            'scopes'                  => 'https://graph.microsoft.com/user.read https://graph.microsoft.com/mail.readwrite https://graph.microsoft.com/mail.send https://graph.microsoft.com/mail.send.shared offline_access'
        ];
    }

}
