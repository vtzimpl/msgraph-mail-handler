<?php 
namespace Vtzimpl\MsgraphMailHandler;

require_once    dirname(__DIR__).'/vendor/autoload.php'; 
use Microsoft\Graph\Graph;
use Microsoft\Graph\Model;
use Microsoft\Graph\Core\GraphConstants;
use GuzzleHttp\Client;
use GuzzleHttp\Exception\BadResponseException;
use GuzzleHttp\Exception\Psr7;
use Microsoft\Graph\Core\ExceptionWrapper;
use Microsoft\Graph\Exception\GraphException;
use Dotenv;


class Gquery
{

private $clientId;
private $clientSecret;
private $tenantId;
private $dotenv;









function __construct() {
$this->dotenv = Dotenv\Dotenv::createImmutable( dirname(__DIR__));
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
	
	
	
	
	
	
	    public function querymOR($userId = null)
    {
		
		

                $InboxFolderID='AAMkADdkNGQ1YjBhLTUwMjYtNDRmNS1iMjY5LTczNTA4ODcwMzY4OQAuAAAAAADI0mtx64SsR5nEA2SqXWvvAQCwi2yxLzSXSLpiqM7jYfQTAAAAAAEMAAA='; //Inbox
				$SentFolderID='AAMkADdkNGQ1YjBhLTUwMjYtNDRmNS1iMjY5LTczNTA4ODcwMzY4OQAuAAAAAADI0mtx64SsR5nEA2SqXWvvAQCwi2yxLzSXSLpiqM7jYfQTAAAAAAEIAAA='; //Sent
				$messageid='AAMkADdkNGQ1YjBhLTUwMjYtNDRmNS1iMjY5LTczNTA4ODcwMzY4OQBGAAAAAADI0mtx64SsR5nEA2SqXWvvBwCwi2yxLzSXSLpiqM7jYfQTAAAAAAEMAACwi2yxLzSXSLpiqM7jYfQTAABcZjWvAAA=';//August Round-up | What's new at Freshworks
                $accessToken = $this->token();
                $graph = new Graph();
                $graph->setAccessToken($accessToken);
        $param=urlencode('"is #114-. "');
                //$user = $graph->createRequest("GET", "/users/51c116ee-aae8-4ec5-9c20-ee0b947f06e1/messages")
				//$messageCollection = $graph->createRequest("GET", "/users/$userId/messages?\$top=30")
				//$messageCollection = $graph->createRequest("GET", "/users/$folderId/messages?\$filter=subject eq 'Automated mail with the deliveries until 2022-08-30 00:00:00'")
				//$messageCollection = $graph->createRequest("GET", "/users/$folderId/messages?\$search=$param")
				//$messageCollection = $graph->createRequest("GET", "/users/$userId/mailFolders")// list mailfolders
				//$messageCollection = $graph->createRequest("GET", "/users/$userId/mailFolders/$InboxFolderID/messages")
				$messageCollection = $graph->createRequest("POST", "/users/$userId/mailFolders/$SentFolderID/messages/$messageid/move")
                               ->setReturnType(Model\Message::class)//->setReturnType(Model\User::class)

                              ->execute();
  
				foreach($messageCollection as $item) {
					//print_r($item->getProperties()['subject']);
					print_r($item);
					print_r("<br>");
														}
	
        return "";
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
	






















	
	
	
		    public function querym2($folderId = null)
    {

                $accessToken = $this->token();
                $graph = new Graph();
                $graph->setAccessToken($accessToken);
        
		
		
		        $messageIterator = $graph->createRequest("GET", "/users/$folderId/messages")
                                         ->setReturnType(Model\Message::class);
										 
        $messages = $messageIterator->getPage();

        while (!$messageIterator->isEnd())
        {
            $messageCollection = $messageIterator->getPage();
			
			
							foreach($messageCollection as $item) {
					print_r($item->getProperties()['subject']);
					print_r("<br>");
														}
			
			
			
			
			
			
			
			
			
        }
     
		
		
		
		
		
		
		
		
                //$user = $graph->createRequest("GET", "/users/51c116ee-aae8-4ec5-9c20-ee0b947f06e1/messages")
				$messageCollection = $graph->createRequest("GET", "/users/$folderId/messages")
                             ->setReturnType(Model\Message::class)//->setReturnType(Model\User::class)
                              ->execute();
  
				foreach($messageCollection as $item) {
					print_r($item->getProperties()['subject']);
					print_r("<br>");
														}
	
        return "";
    }
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	       //get messages from folderId
       // $emails = MsGraph::get("me/mailFolders/$folderId/messages?".$params);
	
	
	
	
	
	
	
	
	
	
	
	
	
}