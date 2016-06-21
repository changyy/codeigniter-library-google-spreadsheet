<?php

set_include_path( get_include_path()
	. PATH_SEPARATOR . dirname(__FILE__) . DIRECTORY_SEPARATOR . 'php-google-spreadsheet-client/src' 
	//. PATH_SEPARATOR . APPPATH. DIRECTORY_SEPARATOR . 'libraries' . DIRECTORY_SEPARATOR .'php-google-spreadsheet-client/src' 
	//. PATH_SEPARATOR . APPPATH. DIRECTORY_SEPARATOR . 'libraries' . DIRECTORY_SEPARATOR . 'google-api-php-client/src' 
);

spl_autoload_register(
	function($className) {
		$className = str_replace("_", "\\", $className);
		$className = ltrim($className, '\\');
		$fileName = '';
		$namespace = '';
        
		if ($lastNsPos = strripos($className, '\\')) {
			$namespace = substr($className, 0, $lastNsPos);
			$className = substr($className, $lastNsPos + 1);
			$fileName = str_replace('\\', DIRECTORY_SEPARATOR, $namespace) . DIRECTORY_SEPARATOR;
		}
		$fileName .= str_replace('_', DIRECTORY_SEPARATOR, $className) . '.php';
		require $fileName;
	}
);

require_once dirname(__FILE__) . DIRECTORY_SEPARATOR . 'vendor/autoload.php';
use Google\Spreadsheet\DefaultServiceRequest;
use Google\Spreadsheet\ServiceRequestFactory;

class Google_Spreadsheet {
	public function reset() {
		$this->client_id = false;
		$this->client_secret_key = false;
		$this->access_token = false;
		$this->spreadsheet = false;
		$this->worksheet = false;
	}

	public function init($access_token, $refresh_token, $client_id, $client_secret_key) {
		$this->reset();

		if (!$this->oauth_check_token($access_token) && false == ($access_token = $this->oauth_renew_access_token($client_id , $client_secret_key, $refresh_token)) ) {
			return 'ACCESS_TOKEN ERROR';
		}

		$this->client_id = $client_id;
		$this->client_secret_key = $client_secret_key;
		$this->access_token = $access_token;
		$this->refresh_token = $refresh_token;

		return true;
	}

	public function find_spreadsheet($name) {
		$this->spreadsheet = false;
		if (!isset($this->access_token))
			return false;

		$serviceRequest = new DefaultServiceRequest(
			$this->access_token
		);
		ServiceRequestFactory::setInstance($serviceRequest);

		$spreadsheetService = new Google\Spreadsheet\SpreadsheetService();
		$spreadsheetFeed = $spreadsheetService->getSpreadsheetFeed();
		try {
			$this->spreadsheet = $spreadsheetFeed->getByTitle($name);
		} catch (Exception $e) {
			$this->spreadsheet = false;
			return $e;
		}

		return $this->spreadsheet !== false;
	}

	public function find_worksheet($name, $auto_create = false, $delete_if_exists = false) {
		if ($this->spreadsheet === false)
			return false;
		$worksheet = false;

		// check exists
		try {
			$worksheetFeed = $this->spreadsheet->getWorksheetFeed();
			$worksheet = $worksheetFeed->getByTitle($name);
			if ($delete_if_exists) {
				$worksheet->delete();
				if ($auto_create)
					$this->spreadsheet->addWorksheet($name);
			}
		} catch (Exception $e) {
			if ($auto_create) {
				$this->spreadsheet->addWorksheet($name);
			} else {
				return false;
			}
		}
		
		try {
			$worksheetFeed = $this->spreadsheet->getWorksheetFeed();
			$worksheet = $worksheetFeed->getByTitle($name);
		} catch (Exception $e) {
			return false;
		}

		$this->worksheet = $worksheet;

		return $this->worksheet !== false;
	}

	public function update_cell($row, $column, $value) {
		return $this->update_cells( array(array($row, $column, $value)) );
	}

	public function update_cells($data = array()) {
		if ($this->worksheet === false || !is_array($data))
			return false;

		$cellFeed = $this->worksheet->getCellFeed();

		try {
			$batchRequest = new Google\Spreadsheet\Batch\BatchRequest();
			foreach($data as $item) {
				if (is_array($item) && count($item) >= 3)
					$batchRequest->addEntry($cellFeed->createCell($item[0], $item[1], $item[2]));
			}
			$batchResponse = $cellFeed->insertBatch($batchRequest);
		} catch (Exception $e) {
			return $e;
		}
		return true;
	}

	public function oauth_request_handler($client_id, $client_secret_key, $scope, $offline = true, $redirect = NULL, $state = NULL) {
		$output = array('action' => 'none');
		$client = new Google_Client();
		$client->setClientId($client_id);
		$client->setClientSecret($client_secret_key);
		// https://developers.google.com/sheets/guides/authorizing#OAuth2Authorizing
		$client->setScopes($scope);
		if ($offline)
			$client->setAccessType('offline');
		if (empty($redirect))
			$redirect = 'http://' . $_SERVER['HTTP_HOST'] . $_SERVER['PHP_SELF'];
		$redirect = filter_var($redirect, FILTER_SANITIZE_URL);
		$client->setRedirectUri($redirect);
		if (!empty($state)) {
			$output['state'] = $state;
			if (!is_array($state))
				$state = http_build_query($state);
			$client->setState($state);
		}
	
		try{
			if (isset($_GET['code'])) {
				$client->authenticate($_GET['code']);
				if (!empty($client->getAccessToken())) {
					$output['action'] = 'done';
					$output['data'] = $client->getAccessToken();
					return $output;
				}
			}
		} catch (Exception $e) {
		}

		if (!isset($_GET['error'])) {
			$output['action'] = 'redirect';
			$output['data'] = $client->createAuthUrl();
		} else {
			$output['action'] = 'error';
			$output['data'] = $_GET['error'];
		}
		return $output;
	}

	function oauth_check_token($token, $min_expire_time = 600) {
		$ret = @json_decode(@file_get_contents(
			'https://www.googleapis.com/oauth2/v1/tokeninfo' . '?' . http_build_query( array(
				'access_token' => $token
			))
		));
		return isset($ret->expires_in) && $ret->expires_in > $min_expire_time;
	}

	function oauth_renew_access_token($client_id, $client_secret_key, $refresh_token) {
		$ch = curl_init();
		curl_setopt($ch, CURLOPT_URL, 'https://www.googleapis.com/oauth2/v3/token');
		curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
		curl_setopt($ch, CURLOPT_POSTFIELDS, http_build_query(array(
			'refresh_token' => $refresh_token,
			'client_id' => $client_id,
			'client_secret' => $client_secret_key,
			'grant_type' => 'refresh_token',
		))); 
		$ret = @json_decode(curl_exec($ch));
		curl_close($ch);
		if (isset($ret->access_token))
			return $ret->access_token;
		return false;
	}
}
