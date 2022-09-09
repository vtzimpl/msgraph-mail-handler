<?php 
namespace Vtzimpl\MsgraphMailHandler;


use Microsoft\Graph\Graph;
use Microsoft\Graph\Model;
use Microsoft\Graph\Core\GraphConstants;
use GuzzleHttp\Client;
use GuzzleHttp\Exception\BadResponseException;
use GuzzleHttp\Exception\Psr7;
use Microsoft\Graph\Core\ExceptionWrapper;
use Microsoft\Graph\Exception\GraphException;
use Dotenv;


class MsgraphMailHandler

{

private $clientId;
private $clientSecret;
private $tenantId;
private $dotenv;









function __construct() {
$this->dotenv = Dotenv\Dotenv::createImmutable(dirname(__DIR__));
$this->dotenv->load();	  
$this->clientId=$_ENV['CLIENT_ID'];
$this->clientSecret=$_ENV['CLIENT_SECRET'];
$this->tenantId=$_ENV['TENAND_ID'];






  }

	
	
 public function token()
    {
        $guzzle = new \GuzzleHttp\Client();
        $url = 'https://login.microsoftonline.com/' . $this->tenantId . '/oauth2/v2.0/token';
        $token = json_decode($guzzle->post($url, [
            'form_params' => [
                'client_id' => $this->clientId,
                'client_secret' => $this->clientSecret,
                'scope' => 'https://graph.microsoft.com/.default',
                'grant_type' => 'client_credentials',
            ],
        ])->getBody()->getContents());
        $accessToken = $token->access_token;
       return $accessToken;
    }
	
	
	//**********************do not delete
				//$messageCollection = $graph->createRequest("GET", "/users/$userId/messages?\$top=30")
				//$messageCollection = $graph->createRequest("GET", "/users/$folderId/messages?\$filter=subject eq 'Automated mail with the deliveries until 2022-08-30 00:00:00'")
				//$messageCollection = $graph->createRequest("GET", "/users/$folderId/messages?\$search=$param")
				//$messageCollection = $graph->createRequest("GET", "/users/$userId/mailFolders")// list mailfolders
				//$messageCollection = $graph->createRequest("GET", "/users/$userId/mailFolders/$InboxFolderID/messages")
				//print_r($item->getProperties()['subject']);
				//->setReturnType(Model\User::class)
	//**********************do not delete
	
	
	
			    public function searchm($userId,$search_term) //general
    {
				$search_term_enc=json_encode($search_term);
                $accessToken = $this->token();
                $graph = new Graph();
                $graph->setAccessToken($accessToken);
				$messageCollection = $graph->createRequest("GET", "/users/$userId//messages?\$search=$search_term_enc")
                                ->setReturnType(Model\Message::class)//->setReturnType(Model\User::class)
                              ->execute();
        return $messageCollection;
    }
	
	
	
	
	
				public function searchms_limitations($userId,$search_term) //limitations toRecipients,subject
    {
				$search_term_enc=json_encode($search_term);
                $accessToken = $this->token();
                $graph = new Graph();
                $graph->setAccessToken($accessToken);
				$messageCollection = $graph->createRequest("GET", "/users/$userId//messages?\$search=$search_term_enc&\$select=toRecipients,subject")
                                ->setReturnType(Model\Message::class)//->setReturnType(Model\User::class)
                              ->execute();
         return $messageCollection;
    }
	
	
					    public function searchms($userId,$search_term) //subject
    {
				$search_term_enc=json_encode($search_term);
                $accessToken = $this->token();
                $graph = new Graph();
                $graph->setAccessToken($accessToken);
				$messageCollection = $graph->createRequest("GET", "/users/$userId//messages?\$search=subject:case")
                                ->setReturnType(Model\Message::class)//->setReturnType(Model\User::class)
                              ->execute();
  

	
        return $messageCollection;
    }
	
	
		    public function movem($userId,$SourceFolderId,$DestinationsFolderId,$MessageId)
    {
                $accessToken = $this->token();
                $graph = new Graph();
                $graph->setAccessToken($accessToken);
                $bodyreq['destinationId']=$DestinationsFolderId;
                $bodyreqj=json_encode($bodyreq);
        		$messageCollection = $graph->createRequest("POST", "/users/$userId/mailFolders/$SourceFolderId/messages/$MessageId/move")
                           ->attachBody($bodyreqj)
                           ->execute();
          return $messageCollection;
    }
	
	
	
	
    public function listm($userId = null)
    {
        $InboxFolderID='AAMkADdkNGQ1YjBhLTUwMjYtNDRmNS1iMjY5LTczNTA4ODcwMzY4OQAuAAAAAADI0mtx64SsR5nEA2SqXWvvAQCwi2yxLzSXSLpiqM7jYfQTAAAAAAEMAAA='; //Inbox
                $accessToken = $this->token();
                $graph = new Graph();
                $graph->setAccessToken($accessToken);

				$messageCollection = $graph->createRequest("GET", "/users/$userId/mailFolders/$InboxFolderID/messages?\$top=30")
                                ->setReturnType(Model\Message::class)//->setReturnType(Model\User::class)
                              ->execute();
  
				foreach($messageCollection as $item) {
					print_r($item->getProperties()['subject']);
                    print_r("<br>");
                    print_r($item->getProperties()['body']);
					
					print_r("<br>");
                    print_r("<br>");
                    print_r("<br>");
														}
	
        return $messageCollection;
    }
	
	
	




    public function listf($userId = null)
    {
      
                $accessToken = $this->token();
                $graph = new Graph();
                $graph->setAccessToken($accessToken);

				$messageCollection = $graph->createRequest("GET", "/users/$userId/mailFolders")// list mailfolders
                                ->setReturnType(Model\Message::class)//->setReturnType(Model\User::class)
                              ->execute();
  
				foreach($messageCollection as $item) {
					//print_r($item->getProperties()['subject']);
					print_r($item);
					print_r("<br>");
														}
	
        return $messageCollection;
    }
	

	
	
	
	
	
	
}
