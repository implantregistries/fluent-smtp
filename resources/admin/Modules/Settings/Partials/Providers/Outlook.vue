<template>
    <div>
        <h3>{{ $t('Outlook / Office365 API Settings') }}</h3>
        <p>{{ $t('Please ') }}<a target="_blank" rel="nofollow" href="https://fluentsmtp.com/docs/setup-outlook-with-fluentsmtp/">{{ $t('check the documentation first to create API keys at Microsoft') }}</a></p>
        <el-form-item>
            <label>{{ $t('Microsoft authentication mode') }}</label>
            <el-radio-group size="mini" v-model="connection.auth_mode">
                <el-radio-button label="delegated">{{ $t('Delegated user authorization') }}</el-radio-button>
                <el-radio-button label="app_only">{{ $t('Application credentials (app-only)') }}</el-radio-button>
            </el-radio-group>
            <p v-if="connection.auth_mode === 'app_only'">
                {{ $t('Create and authorize the Microsoft application in Microsoft 365, then enter its credentials below. Exchange Online Application RBAC is recommended to scope mailbox access; this connector does not restrict tenant permissions.') }}
            </p>
            <p v-if="connection.auth_mode === 'app_only'">
                {{ $t('Microsoft Graph uses the Exchange mailbox display name for delivered app-only messages and may replace the From Name configured above. Set the mailbox display name in Microsoft 365 when a specific delivered name is required.') }}
            </p>
        </el-form-item>
        <el-radio-group size="mini" v-model="connection.key_store">
            <el-radio-button value="db" label="db">{{ $t('Store application credentials in FluentSMTP (encrypted)') }}</el-radio-button>
            <el-radio-button value="wp_config" label="wp_config">{{ $t('Application Keys in Config File') }}</el-radio-button>
        </el-radio-group>

        <p v-if="connection.key_store == 'db' && connection.auth_mode === 'app_only'">
            {{ $t('Enter the Microsoft application credentials below. FluentSMTP stores the client-secret value using its existing encrypted settings workflow.') }}
        </p>

        <el-row :gutter="20" v-if="connection.key_store == 'db'">
            <el-col :md="12" :sm="24" v-if="connection.auth_mode === 'app_only'">
                <el-form-item>
                    <label for="tenant_id">
                        {{ $t('Directory (tenant) ID') }}
                    </label>
                    <el-input id="tenant_id" v-model="connection.tenant_id" />
                    <error :error="errors.get('tenant_id')" />
                </el-form-item>
            </el-col>
            <el-col :md="12" :sm="24">
                <el-form-item>
                    <label for="client_id">
                        {{ $t('Application Client ID') }}
                    </label>

                    <InputPassword
                        id="client_id"
                        v-model="connection.client_id"
                        :disable_help="connection.disable_encryption === 'yes'"
                    />

                    <error :error="errors.get('client_id')" />
                </el-form-item>
            </el-col>

            <el-col :md="12" :sm="24">
                <el-form-item>
                    <label for="client_secret">
                        {{ $t('Application Client Secret Value (not the Secret ID)') }}
                    </label>

                    <InputPassword
                        id="client_secret"
                        v-model="connection.client_secret"
                        :disable_help="connection.disable_encryption === 'yes'"
                    />
                    <error :error="errors.get('client_secret')" />
                </el-form-item>
            </el-col>

            <el-col :md="24">
                <el-form-item>
                    <el-checkbox true-label="yes" false-label="no" v-model="connection.disable_encryption">
                        {{ $t('Disable Encryption for Application Client Secret (Not Recommended)') }}
                    </el-checkbox>
                    <p style="color: red; margin-top: 0;" v-if="connection.disable_encryption === 'yes'">
                        {{
                            $t('By disabling encryption, your Application Client Secret will be stored in plain text in the database. This is not recommended for security reasons. Enable only if your security plugin rotate WP SALTS frequently.')
                        }}
                    </p>
                </el-form-item>
            </el-col>

        </el-row>

        <div class="fss_condesnippet_wrapper" v-else-if="connection.key_store == 'wp_config'">
            <el-form-item>
                <label>{{ $t('__WP_CONFIG_INSTRUCTION') }}</label>
                <div class="code_snippet">
                    <textarea v-if="connection.auth_mode === 'app_only'" readonly style="width: 100%;">define( 'FLUENTMAIL_OUTLOOK_CLIENT_ID', '********************' );
define( 'FLUENTMAIL_OUTLOOK_CLIENT_SECRET', '********************' );
define( 'FLUENTMAIL_OUTLOOK_TENANT_ID', '********************' );
define( 'FLUENTMAIL_OUTLOOK_AUTH_MODE', 'app_only' );</textarea>
                    <textarea v-else readonly style="width: 100%;">define( 'FLUENTMAIL_OUTLOOK_CLIENT_ID', '********************' );
define( 'FLUENTMAIL_OUTLOOK_CLIENT_SECRET', '********************' );</textarea>
                </div>
                <error :error="errors.get('client_id')" />
                <error :error="errors.get('client_secret')" />
                <error v-if="connection.auth_mode === 'app_only'" :error="errors.get('tenant_id')" />
            </el-form-item>
        </div>

        <el-form-item v-if="connection.auth_mode === 'delegated'">
            <label>{{ $t('App Callback URL(Use this URL to your APP)') }}</label>
            <el-input :readonly="true" v-model="provider.callback_url" />
        </el-form-item>

        <div v-if="connection.auth_mode === 'delegated' && !connection.access_token">
            <div style="text-align: center;">
                <h3>{{ $t('Please authenticate with Office365 to get ') }}<b>{{ $t('Access Token') }}</b></h3>
                <el-button v-loading="gettingRedirect" @click="redirectToMS()" type="danger">{{ $t('Authenticate with Office365 & Get Access Token') }}</el-button>
            </div>
            <el-row v-if="redirectUrl" :gutter="20">
                <el-col :span="12">
                    <el-form-item>
                        <label for="application_token">
                            {{ $t('Access Token') }}
                        </label>
                        <InputPassword
                            id="application_token"
                            v-model="connection.auth_token"
                        />
                        <error :error="errors.get('auth_token')" />
                        <p>{{ $t('Please send test email to confirm if the connection is working or not.') }}</p>
                    </el-form-item>
                </el-col>
            </el-row>
        </div>
        <div style="text-align: center;" v-else-if="connection.auth_mode === 'delegated'">
            <h3>{{ ('Your Outlook / Office365 Authentication has been enabled.No further action is needed.If you want to re-authenticate, ') }}<a @click.prevent="connection.access_token = ''" href="#">{{ ('click here') }}</a></h3>
        </div>
        <div v-else>
            <p>{{ $t('No interactive Microsoft login or callback is required. FluentSMTP will obtain and cache an application access token when the connection is saved.') }}</p>
        </div>

    </div>
</template>

<script type="text/babel">
    import InputPassword from '@/Pieces/InputPassword';
    import Error from '@/Pieces/Error';

    export default {
        name: 'OutLook',
        props: ['connection', 'provider', 'errors'],
        components: {
            InputPassword,
            Error
        },
        data() {
            return {
                app_ready: false,
                gettingRedirect: false,
                redirectUrl: ''
            };
        },
        watch: {
            'connection.auth_mode'(value, oldValue) {
                if (oldValue && value !== oldValue) {
                    this.connection.auth_token = '';
                    this.connection.access_token = '';
                    this.connection.refresh_token = '';
                    this.connection.expire_stamp = 0;
                }
                if (value === 'delegated') {
                    this.connection.tenant_id = '';
                }
            },
            'connection.key_store'(value) {
                if (value === 'wp_config') {
                    this.connection.client_id = '';
                    this.connection.client_secret = '';
                    this.connection.tenant_id = '';
                }
            }
        },
        methods: {
            redirectToMS() {
                this.gettingRedirect = true;
                this.$post('settings/outlook_auth_url', {
                    connection: this.connection
                })
                    .then(response => {
                        this.redirectUrl = response.data.auth_url;
                        window.open(response.data.auth_url, '_blank');
                    })
                    .catch(errors => {
                        this.errors.record(errors.responseJSON.data);
                    })
                    .always(() => {
                        this.gettingRedirect = false;
                    });
            }
        },
        mounted() {
            if (!this.connection.auth_mode) {
                this.$set(this.connection, 'auth_mode', 'delegated');
            }
            if (!this.connection.key_store) {
                this.$set(this.connection, 'key_store', 'db');
            }
        }
    };
</script>
