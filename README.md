# Init

Step 1: 

```
$ git clone https://github.com/changyy/codeigniter-library-google-spreadsheet.git
```

Step 2:

```
$ cd /tmp && curl -sS https://getcomposer.org/installer | php && cd -
$ php /tmp/composer.phar install
```

# Basic Usage

```
<?php
$obj = new Google_Spreadsheet();

if (true === $obj->init(
	'user_access_token',
	'user_refresh_token',
	'client_id',
	'client_secret_key'
)) {
	if (true === $this->google_spreadsheet->find_spreadsheet('MySpreadsheetsName', true)) {	// true for auto create
		if (true === $this->google_spreadsheet->find_worksheet('MyWorksheet', true)) {	// true for auto create
			// single update: row, column, value
			if (true === $this->google_spreadsheet->update_cell(1, 1, date('Y-m-d H:i:s'))) {
				echo "Insert\n";
			}

			// batch update
			if (true === $this->google_spreadsheet->update_cells(array(
				array( 1, 1, date('Y-m-d H:i:s') ),
				array( 1, 2, 'Hello' ),
				array( 1, 3, 'World' ),
			))) {
				echo "Batch mode insert\n";
			}

			if (true === $this->google_spreadsheet->frozen_current_worksheet(1, 0)) {
				echo "setFrozen (Row, Column) = (1, 0)\n";
			}
		}
	}
}
```

# CodeIgniter Usage

## Install

```
$ cd /path/ci/project/application/libraries
$ git submodule add https://github.com/changyy/codeigniter-library-google-spreadsheet.git
$ php composer.phar install
```

## CodeIgniter Controller

```
$ vim /path/ci/project/application/controllers/Welcome.php
<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Welcome extends CI_Controller {
	public function index() {
		$this->load->library('codeigniter-library-google-spreadsheet/Google_Spreadsheet');
		$this->google_spreadsheet->init(
			'user_access_token',
			'user_refresh_token',
			'client_id',
			'client_secret_key'
		);

		if (true === $this->google_spreadsheet->find_spreadsheet('MySpreadsheetsName')) {
			echo "pass find_spreadsheet \n";
			if (true === $this->google_spreadsheet->find_worksheet('MyWorksheet', true)) {	// true: create if not exists
				echo "pass find_worksheet \n";
				if (true === $this->google_spreadsheet->update_cell(1, 1, date('Y-m-d H:i:s'))) {
					echo "pass update_cell \n";
				}

				if (true === $this->google_spreadsheet->update_cells(array(
					array( 1, 1, date('Y-m-d H:i:s') ),
					array( 1, 2, 'Hello' ),
					array( 1, 3, 'World' ),
					
				))) {
					echo "pass update_cells \n";
				}
			}
		}
	}

	public function oauth() {
		$this->load->helper('url');
		$this->load->library('codeigniter-library-google-spreadsheet/Google_Spreadsheet');
		$ret = $this->google_spreadsheet->oauth_request_handler(
			'client_id',
			'client_secret_key'
			// scope
			array(
				'https://www.googleapis.com/auth/spreadsheets', 
				'https://spreadsheets.google.com/feeds',
			),
			// enable offline
			true,
			preg_replace("{//[^/]+/}", "//".$this->input->server('HTTP_HOST')."/", current_url())
		);

		if ($ret['action'] == 'redirect') {
			redirect($ret['data'], 'location', 301);
			return;
		}

		print_r($ret);
		// Array ( [action] => done [data] => Array ( [access_token] => UsersAceessToken [token_type] => Bearer [expires_in] => 3600 [created] => 1466331600 ) )
	}
}
```

## Setup the Google cloud platform & get User's access token and refresh token

- https://console.cloud.google.com/apis/library
  - API Mamager
    - Overview
      - enable Sheets API
    - Credentails
      - Create credentials -> OAuth 2.0 client ID -> Web application
        - Authorized redirect URIs
          - Use CodeIgniter Project path, like: https://localhost/ci/project/index.php/welcome/oauth
