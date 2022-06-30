<?php

namespace App\Controller\Component;

use Cake\Core\Configure;
use Cake\Controller\Component;
use Exception;
use Cake\Http\Exception\BadRequestException;

use Cake\Filesystem\Folder;
use Cake\Filesystem\File;


use Google_Client;
use Google_Service_Exception;
use Google_Service_Sheets;
use Google_Service_Sheets_Spreadsheet;
use Google_Service_Sheets_Request;
use Google_Service_Sheets_BatchUpdateSpreadsheetRequest;
use Google_Service_Sheets_ValueRange;
use Google_Service_Drive;
use Google_Service_Drive_Permission;

class GoogleSheetComponent extends Component
{
    public function initialize(array $config)
    {
        parent::initialize($config);
    }

    public function getClient()
    {
        $credentials_path = CONFIG."oauth_credentials.json";
        $client = new Google_Client();
        $client->setAuthConfig($credentials_path);
        $client->setScopes([
            Google_Service_Sheets::DRIVE_FILE, // ドライブ
        ]);

        return $client;
    }

    public function oauthRedirectUri()
    {
        $client = $this->getClient();

        // 認証後のリダイレクトURLを設定
        $client->setRedirectUri('https://' . $_SERVER['HTTP_HOST'] . 'hogehoge');

        // 認証画面を設定
        $client->setPrompt("consent");

        // 偽造防止のためにトークンを指定
        $state = bin2hex(random_bytes(128/8));
        $session = $this->request->session();
        $session->write('google.stateToken', $state);
        $client->setState($state);

        // GoogleのOAuth2.0サーバーへリクエストを行うためのURLを生成する
        $auth_url = $client->createAuthUrl();
        header('Location: ' . filter_var($auth_url, FILTER_SANITIZE_URL));

    }

    public function outputSheet($spreadName,$spreadData)
    {
        $client = $this->getClient();

        // ユーザー認証で返ってきたトークンを設定
        $client->authenticate($_GET['code']);
        $accessToken = $client->getAccessToken();
        $client->setAccessToken($accessToken);

        try {
            // マイドライブ内を検索
            // 条件：ゴミ箱フォルダを除いたスプレッドシートの指定ファイル名
            $drive_service = new Google_Service_Drive($client);
            $result = $drive_service->files->listFiles([
                "q" => "name ='$spreadName' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false",
            ]);
        } catch (Google_Service_Exception $exception) {
            if (Configure::read('debug')) {
                throw $exception;
            }
            throw new BadRequestException(__('GoogleDriveでの認証が失敗しました。'));
        }

        //// スプレッドシートの処理
        $spreadsheet_service = new Google_Service_Sheets($client);

        $sheet_name = date("Y年m月d日H時i分s秒");
        $spread_id = '';
        $sheet_id = '';

        // 新規か更新の判定
        if(empty($result["files"][0]['id'])){
            // スプレットシート新規作成
            $requestBody = new Google_Service_Sheets_Spreadsheet([
                'properties' => [
                    'title' => $spreadName
                ],
                "sheets" => [
                    'properties' => [
                        'title' => $sheet_name
                    ],
                ],
            ]);

            try {
                $response = $spreadsheet_service->spreadsheets->create($requestBody);
            } catch (Google_Service_Exception $exception) {
                if (Configure::read('debug')) {
                    throw $exception;
                }
                throw new BadRequestException(__('SpreadSheetの作成に失敗しました。'));
            }

            $spread_id = $response->spreadsheetId;
            $sheet_id = $response->sheets[0]->properties->sheetId;

        }else{
            // 更新
            $spread_id = $result["files"][0]['id'];

            // シートの追加
            $body = new \Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
                'requests' => [
                    'addSheet' => [
                        'properties' => [
                            'title' => $sheet_name
                        ]
                    ]
                ]
            ]);
            try {
                $response = $spreadsheet_service->spreadsheets->batchUpdate($spread_id, $body);
            } catch (Google_Service_Exception $exception) {
                if (Configure::read('debug')) {
                    throw $exception;
                }
                throw new BadRequestException(__('SpreadSheetのシート追加に失敗しました。'));
            }

            $sheet_id = $response->getReplies()[0]
            ->getAddSheet()
            ->getProperties()
            ->sheetId;
        }

        // シートへデータ入力
        $sheet_range = "A1"; // 範囲を指定も出来ます。A1:C5
        $update = new Google_Service_Sheets_ValueRange([
            'values' => $spreadData,
        ]);

        try {
            $response = $spreadsheet_service->spreadsheets_values->update(
                $spread_id,
                $sheet_name.'!'.$sheet_range,
                $update,
                ["valueInputOption" => 'USER_ENTERED']
            );
        } catch (Google_Service_Exception $exception) {
            if (Configure::read('debug')) {
                throw $exception;
            }
            throw new BadRequestException(__('SpreadSheetのデータ入力に失敗しました。'));
        }


        // シートデザイン調整
        $body = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
            'requests' => $this->createRequestData($sheet_id,$spreadData)
        ]);

        try {
            $response = $spreadsheet_service->spreadsheets->batchUpdate(
                $spread_id,
                $body
            );
        } catch (Google_Service_Exception $exception) {
            if (Configure::read('debug')) {
                throw $exception;
            }
            throw new BadRequestException(__('SpreadSheetのレイアウト調整に失敗しました。'));
        }

        $spreadSheet_url = 'https://docs.google.com/spreadsheets/d/'.$spread_id.'/edit#gid='.$sheet_id;

        return $spreadSheet_url;
    }

    public function createRequestData($sheet_id, $spreadData)
    {
        $endRow = count($spreadData);
        $endColumn = count($spreadData[0]);

        $request_data = [];

        // 罫線の設定
        $border = [
            "style"=> "SOLID",
            "width"=> 1,
            "color"=> [
                "red"=> 0,
                "green"=> 0,
                "blue"=> 0,
                "alpha"=> 0
            ]
        ];

        $border_data = new Google_Service_Sheets_Request([
            'updateBorders' => [
                'range' => [
                    'sheetId' => $sheet_id,
                    'startRowIndex' => 0,       // 行の開始位置
                    'endRowIndex' => $endRow,         // 行の終了位置
                    'startColumnIndex' => 0,    // 列の開始位置
                    'endColumnIndex' => $endColumn,      // 列の終了位置
                ],
                'top' => $border,
                'bottom' => $border,
                'right' => $border,
                'left' => $border,
                "innerHorizontal" => $border,
                "innerVertical" => $border
            ]
        ]);

        // 自動列幅調整の設定
        $colum_resize_data = new Google_Service_Sheets_Request([
            "autoResizeDimensions"=> [
                "dimensions"=> [
                    "sheetId"=> $sheet_id,
                    "dimension"=> "COLUMNS",
                    "startIndex"=> 0,
                    "endIndex"=> $endColumn
                ]
            ]
        ]);

        // 「交互の背景色」の設定
        $banding_data = new Google_Service_Sheets_Request([
            "addBanding"=>[
                "bandedRange" => [
                    'range' => [
                        'sheetId' => $sheet_id,
                        'startRowIndex' => 0,
                        'endRowIndex' => $endRow,
                        'startColumnIndex' => 0,
                        'endColumnIndex' => $endColumn,
                    ],
                    "rowProperties" => [
                        "headerColor" => [
                            "red"=> 189/255,
                            "green"=> 189/255,
                            "blue"=> 189/255,
                            "alpha"=> 1/255
                        ],
                        "firstBandColor" => [
                            "red"=> 255/255,
                            "green"=> 255/255,
                            "blue"=> 255/255,
                            "alpha"=> 0/255
                        ],
                        "secondBandColor" => [
                            "red"=> 243/255,
                            "green"=> 243/255,
                            "blue"=> 243/255,
                            "alpha"=> 0/255
                        ]
                    ]
                ]
            ]
        ]);

        $request_data = array($border_data, $colum_resize_data, $banding_data);

        return $request_data;
    }
}
