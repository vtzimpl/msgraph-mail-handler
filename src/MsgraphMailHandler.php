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
$this->dotenv = Dotenv\Dotenv::createImmutable(base_path());
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
                    $separator="Reply above this line.";
                    $return_array=array();
                    $search_term_enc=json_encode($search_term);
                    $accessToken = $this->token();
                    $graph = new Graph();
                    $graph->setAccessToken($accessToken);
                    $messageCollection = $graph->createRequest("GET", "/users/$userId//messages?\$search=$search_term_enc")
                                            ->setReturnType(Model\Message::class)//->setReturnType(Model\User::class)
                                          ->execute();

                                         foreach($messageCollection as $message)
                                         {
                                             $mes_part['subject']=$message->getProperties()['subject'];
                                             //$mes_part['body']=$message->getProperties()['body']['content'];
                                             $mes_part['body']=explode($separator,htmlspecialchars(trim(strip_tags($message->getProperties()['body']['content']))))[0];
                                             array_push($return_array, $mes_part);
                                         }
                                         
                                          
                                          
                                          
                                     // dd($return_array    );
                                          
                                          
                                          
                                          
                    return $return_array;
                }
	
               
               
               
               
                public function searchmf($userId,$searchinfolderid,$search_term) //search in particular folder
                {
                    $separator="Reply above this line.";
                    $return_array=array();
                    $search_term_enc=json_encode($search_term);
                    $accessToken = $this->token();
                    $graph = new Graph();
                    $graph->setAccessToken($accessToken);
                    $messageCollection = $graph->createRequest("GET", "/users/$userId/mailFolders/$searchinfolderid/messages?\$search=$search_term_enc")
                                            ->setReturnType(Model\Message::class)//->setReturnType(Model\User::class)
                                          ->execute();
                                         //dd($messageCollection);
                                         foreach($messageCollection as $message)
                                         {
                                            //dd($message);
                                            //support_case_id, mail_id, mail_internetMessageId, mail_subject, mail_body_content, mail_weblink, mail_conversationId, mail_conversationIndex, mail_from, mail_sentDateTime, mail_receivedDateTime, mail_type, message_type
                                            $mes_part['hasAttachments']=$message->getProperties()['hasAttachments'];
                                            $mes_part['mail_id']=$message->getProperties()['id'];
                                            if($mes_part['hasAttachments']){

                                                $this->listattac($userId,$searchinfolderid,$mes_part['mail_id']);


                                            }
                                            
                                            
                                            
                                            $mes_part['mail_internetMessageId']=$message->getProperties()['internetMessageId'];
                                            $mes_part['mail_subject']=$message->getProperties()['subject'];
                                            $mes_part['mail_weblink']=$message->getProperties()['webLink'];
                                            $mes_part['mail_conversationId']=$message->getProperties()['conversationId'];
                                             $mes_part['mail_conversationIndex']=$message->getProperties()['conversationIndex'];
                                             $mes_part['mail_from']=$message->getProperties()['from']['emailAddress']['address'];
                                             $mes_part['mail_sentDateTime']=$message->getProperties()['sentDateTime'];
                                             $mes_part['mail_receivedDateTime']=$message->getProperties()['receivedDateTime'];
                                             //$mes_part['body']=$message->getProperties()['body']['content'];
                                             $mes_part['mail_body_content']=explode($separator,htmlspecialchars(trim(strip_tags($message->getProperties()['body']['content']))))[0];
                                             array_push($return_array, $mes_part);
                                         }
                                         
                                          
                                          
                                          
                                      //dd($return_array    );
                                          
                                          
                                          
                                          
                    return $return_array;
               
               
               
               
                                        }
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
                
	
	
	
				public function searchms_limitations($userId,$search_term) //limitations toRecipients,subject
    {
				$search_term_enc=json_encode($search_term);
                $accessToken = $this->token();
                $graph = new Graph();
                $graph->setAccessToken($accessToken);
				$messageCollection = $graph->createRequest("GET", "/users/$userId/messages?\$search=$search_term_enc&\$select=toRecipients,subject")
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
				$messageCollection = $graph->createRequest("GET", "/users/$userId/messages?\$search=subject:case")
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
	

	
	
	

    public function listattac($userId,$searchinfolderid,$mail_id) //list attachments
    {
      
                $accessToken = $this->token();
                $graph = new Graph();
                $graph->setAccessToken($accessToken);
     			$messageCollection = $graph->createRequest("GET", "/users/$userId/mailFolders/$searchinfolderid/messages/$mail_id/attachments")// list mailfolders
                                ->setReturnType(Model\Message::class)//->setReturnType(Model\User::class)
                              ->execute();
  //dd($messageCollection);
				foreach($messageCollection as $item) {
					//print_r($item->getProperties()['subject']);
					print_r($item);
					print_r("<br>");
														}
	
        return $messageCollection;
    }





    




	
	
	
}
