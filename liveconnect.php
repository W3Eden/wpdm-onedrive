<?php

function_exists('add_action') or die;

/**
 * Check Authentication.
 */
class LiveConnectClient {

    protected $options;
    protected $http;
    protected $access_id = 0;
    protected static $access = array();
    protected static $token = array();
    protected static $redirect_uri = null;
    protected static $instance = null;

    public function __construct($options = null) {
        $wpdm_onedrive = maybe_unserialize(get_option('__wpdm_onedrive', array()));
        $this->options = array(
            'clientid' => $wpdm_onedrive('client_id'),
            'clientsecret' => $wpdm_onedrive('client_secret'),
            'sendheaders' => false,
            'authmethod' => 'get',
            'authurl' => 'https://login.live.com/oauth20_token.srf',
            'tokenurl' => 'https://login.live.com/oauth20_token.srf',
            'redirecturi' => $wpdm_onedrive('redirect_url'),
//            self::getRedirectUri()
            'userefresh' => true,
            'storetoken' => true,
            'usecookie' => true,
            'cookiename' => 'wl_auth',
            'timeout' => 25,
            'sslverify' => false
        );

        if (is_array($options))
            $this->options = array_merge($this->options, (array) $options);

        $this->http = new WP_Http;
    }

    public static function getInstance($options = null) {
        // Automatically instantiate the singleton object if not already done.
        if (empty(self::$instance)) {
            self::setInstance(new LiveConnectClient($options));
        }
        return self::$instance;
    }

    public static function setInstance($instance) {
        if (($instance instanceof LiveConnectClient) || $instance === null) {
            self::$instance = & $instance;
        }
    }

    public static function getUserIdFromResource($resource_id) {
        if (preg_match('/^.+\.([0-9a-f]+)\..+$/', $resource_id, $match)) {
            return $match[1];
        }
        return null;
    }

    /**
     * Get an option from the LiveConnectClient instance.
     *
     * @param   string  $key  The name of the option to get
     *
     * @return  mixed  The option value
     */
    public function getOption($key) {
        return array_key_exists($key, $this->options) ? $this->options[$key] : null;
    }

    /**
     * Set an option for the LiveConnectClient instance.
     *
     * @param   string  $key    The name of the option to set
     * @param   mixed   $value  The option value to set
     *
     * @return  LiveConnectClient  This object for method chaining
     */
    public function setOption($key, $value) {
        $this->options[$key] = $value;

        return $this;
    }

    /**
     * Get the access token or redict to the authentication URL.
     *
     * @return  WP_Error|string The access token or WP_Error on failure.
     */
    public function authenticate() {
        $this->log(__METHOD__);

        if (isset($_GET['code']) AND ( $data['code'] = $_GET['code'])) {
            $wpdm_onedrive = maybe_unserialize(get_option('__wpdm_onedrive', array()));
            $data['grant_type'] = 'authorization_code';
            $data['redirect_uri'] = $wpdm_onedrive('redirect_url');
            $data['client_id'] = $wpdm_onedrive('client_id');
            $data['client_secret'] = $wpdm_onedrive('client_secret');

            $response = $this->http->post($this->getOption('tokenurl'), array('body' => $data, 'timeout' => $this->getOption('timeout'), 'sslverify' => $this->getOption('sslverify')));

            if (is_wp_error($response)) {
                return $response;
            } elseif ($response['response']['code'] >= 200 AND $response['response']['code'] < 400) {
                if (isset($response['headers']['content-type']) AND strpos($response['headers']['content-type'], 'application/json') !== false) {
                    $token = array_merge(json_decode($response['body'], true), array('created' => time()));
                }

                $this->setToken($token);

                return $token;
            } else {
                return new WP_Error('oauth_failed', 'Error code ' . $response['response']['code'] . ' received requesting access token: ' . $response['body'] . '.');
            }
        }

        if ($this->getOption('sendheaders')) {
            // If the headers have already been sent we need to send the redirect statement via JavaScript.
            if (headers_sent()) {
                echo "<script>document.location.href='" . str_replace("'", "&apos;", $this->createUrl()) . "';</script>\n";
            } else {
                // All other cases use the more efficient HTTP header for redirection.
                header('HTTP/1.1 303 See other');
                header('Location: ' . $this->createUrl());
                header('Content-Type: text/html; charset=utf-8');
            }

            die();
        }
        return false;
    }

    /**
     * Verify if the client has been authenticated
     *
     * @return  boolean  Is authenticated
     */
    public function isAuthenticated() {
        $this->log(__METHOD__);

        $token = $this->getToken();

        if (!$token || !array_key_exists('access_token', $token)) {
            return false;
        } elseif (array_key_exists('expires_in', $token) && $token['created'] + $token['expires_in'] < time() + 20) {
            return false;
        } else {
            return true;
        }
    }

    /**
     * Create the URL for authentication.
     *
     * @return  WP_Error|string The URL or WP_Error on failure.
     */
    public function createUrl() {
        $this->log(__METHOD__);
        $wpdm_onedrive = maybe_unserialize(get_option('__wpdm_onedrive', array()));

        if (!$this->getOption('authurl') || !$wpdm_onedrive('client_id')) {
            return new WP_Error('oauth_failed', 'Authorization URL and client_id are required');
        }

        $url = $this->getOption('authurl');

        if (strpos($url, '?')) {
            $url .= '&';
        } else {
            $url .= '?';
        }

        $url .= 'response_type=code';
        $url .= '&client_id=' . urlencode($wpdm_onedrive('client_id'));

        if ($this->getOption('redirecturi')) {
            $url .= '&redirect_uri=' . urlencode($wpdm_onedrive('redirect_url'));
        }

        if ($this->getOption('scope')) {
            $scope = is_array($this->getOption('scope')) ? implode(' ', $this->getOption('scope')) : $this->getOption('scope');
            $url .= '&scope=' . urlencode($scope);
        }

        if ($this->getOption('state')) {
            $url .= '&state=' . urlencode($this->getOption('state'));
        }

        if (is_array($this->getOption('requestparams'))) {
            foreach ($this->getOption('requestparams') as $key => $value) {
                $url .= '&' . $key . '=' . urlencode($value);
            }
        }

        return $url;
    }

    public function handlePageRequest() {
        $this->log(__METHOD__);

        if ($this->isAuthenticated() OR ( isset($_GET['access_token']) AND $_GET['access_token'])) {
            $this->log(__METHOD__ . '. There is a token available already');
            // There is a token available already. It should be the token flow. Ignore it.
            return;
        }

        $token = $this->authenticate();
        if (is_wp_error($token)) {
            $this->log(__METHOD__ . '. Authentication error: ' . $token->get_error_message(), E_USER_ERROR);
            $token = false;
        }

        if ($token === false) {
            $token = $this->loadToken();
            if (is_array($token) && array_key_exists('expires_in', $token) && $token['created'] + $token['expires_in'] < time() + 20) {
                if (!$this->getOption('userefresh')) {
                    return false;
                }

                $token = $this->refreshToken($token['refresh_token']);
                if (is_wp_error($token)) {
                    $this->log(__METHOD__ . '. Refreshing token error: ' . $token->get_error_message(), E_USER_ERROR);
                    return false;
                }
            }
        }

        $error = array(
            'error' => isset($_GET['error']) ? $_GET['error'] : null,
            'error_description' => isset($_GET['error_description']) ? $_GET['error_description'] : null
        );

        if ($error['error']) {
            $this->setToken($error);
        }

        $this->log(__METHOD__ . '. End');

        return
                '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">' .
                '<html xmlns="http://www.w3.org/1999/xhtml" xmlns:msgr="http://messenger.live.com/2009/ui-tags">' .
                '<head>' .
                '<title>Live SDK Callback Page</title>' .
                '<script src="//js.live.net/v5.0/wl.js" type="text/javascript"></script>' .
                '</head>' .
                '<body></body>' .
                '</html>';
    }


    public function getToken() {
        $this->log(__METHOD__);

        $token = isset(self::$token[$this->access_id]) ?
                self::$token[$this->access_id] :
                (isset(self::$token[0]) ? self::$token[0] : null);

        if (!$token AND $this->loadToken()) {
            $token = self::$token[$this->access_id];
        }

        return $token;
    }

    public function query($url = null, $data = null, $headers = array(), $method = 'get', $timeout = null) {
        $this->log(__METHOD__ . '. URL: ' . $url . ' ' . print_r($data, true));

        $url = strpos($url, 'http') === 0 ? $url : 'https://apis.live.net/v5.0/' . ltrim($url, '/');

        $token = $this->getToken();
        if (array_key_exists('expires_in', $token) && $token['created'] + $token['expires_in'] < time() + 20) {
            if (!$this->getOption('userefresh')) {
                return false;
            }
            $token = $this->refreshToken($token['refresh_token']);
        }

        if (!$this->getOption('authmethod') || $this->getOption('authmethod') == 'bearer') {
            $headers['Authorization'] = 'Bearer ' . $token['access_token'];
        } elseif ($this->getOption('authmethod') == 'get') {
            if (strpos($url, '?')) {
                $url .= '&';
            } else {
                $url .= '?';
            }
            $url .= $this->getOption('getparam') ? $this->getOption('getparam') : 'access_token';
            $url .= '=' . $token['access_token'];
        }

        $args = array(
            'method' => $method,
            'headers' => $headers,
            'timeout' => $timeout > 0 ? $timeout : $this->getOption('timeout'),
            'sslverify' => $this->getOption('sslverify')
        );

        switch ($method) {
            case 'get':
            case 'delete':
            case 'trace':
            case 'head':
                break;
            case 'post':
            case 'put':
            case 'patch':
                $args['body'] = $data;
                break;
            default:
                return new WP_Error('oauth_failed', 'Unknown HTTP request method: ' . $method . '.');
        }

        $response = $this->http->request($url, $args);

        $this->log(__METHOD__ . '. ' . print_r($response, true));

        if (is_wp_error($response)) {
            $this->log(__METHOD__ . '. Request error: ' . $response->get_error_message(), E_USER_ERROR);
        } elseif ($response['response']['code'] < 200 OR $response['response']['code'] >= 400) {
            $error = __METHOD__ . '. Response code ' . $response['response']['code'] . ' received requesting data: ' . $response['body'] . '.';
            $this->log($error, E_USER_ERROR);
            return new WP_Error('oauth_failed', $error);
        } elseif (isset($response['headers']['content-type']) AND strpos($response['headers']['content-type'], 'application/json') !== false) {
            $response['body'] = json_decode($response['body']);
        }

        return $response;
    }


    public function request($url = null, $headers = array()) {
        $this->log(__METHOD__ . '. URL: ' . $url);

        $response = $this->http->get($url, array('headers' => $headers, 'timeout' => $this->getOption('timeout'), 'sslverify' => $this->getOption('sslverify')));

        if (is_wp_error($response)) {
            $this->log(__METHOD__ . '. Request error ' . $response->get_error_message(), E_USER_ERROR);
        } elseif ($response['response']['code'] < 200 OR $response['response']['code'] >= 400) {
            $error = __METHOD__ . '. Response code ' . $response['response']['code'] . ' received requesting data: ' . $response['body'] . '.';
            $this->log($error, E_USER_ERROR);
            return new WP_Error('request_failed', $error);
        }

        return $response;
    }

    protected function buildQueryString($array) {
        $result = '';
        foreach ($array as $k => $v) {
            if ($result == '') {
                $prefix = '';
            } else {
                $prefix = '&';
            }
            $result .= $prefix . rawurlencode($k) . '=' . rawurlencode($v);
        }

        return $result;
    }

    protected function parseQueryString($query) {
        $result = array();
        $arr = preg_split('/&/', $query);
        foreach ($arr as $arg) {
            if (strpos($arg, '=') !== false) {
                $kv = preg_split('/=/', $arg);
                $result[rawurldecode($kv[0])] = rawurldecode($kv[1]);
            }
        }
        return $result;
    }

    public function log($message = null, $level = E_USER_NOTICE) {
//		if (!PWEB_ONEDRIVE_DEBUG) return;

        switch ($level) {
            case E_USER_ERROR:
                $prefix = '   Error     ';
                break;
            case E_USER_WARNING:
                $prefix = '   Warning   ';
                break;
            case E_USER_NOTICE:
            default:
                $prefix = '   Notice    ';
        }

        error_log("\r\n" . date('Y-m-d H:i:s') . $prefix . $message, 3, WP_CONTENT_DIR . '/debug.log');
    }

}
