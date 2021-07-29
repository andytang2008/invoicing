<?php
include "SimpleXLSX.php";   //It is the file used by code to open and parse the Microsoft Excel file.
ini_set("memory_limit","360M"); //Setup the maximum file size you will open.
global $enterkey, $tabkey;     //    \r (Carriage Return)    \n (Line Feed) 
$enterkey="\r\n";
$tabkey="\t";                 //  \t  tab

if ( $xlsx = SimpleXLSX::parse('invoice_full.xlsx') ) {   //invoice_full_xlsx is the file containing invoice and invoice line data.
    $i = 0;

    foreach ($xlsx->rows() as $elt) {   //$elt is the array containing the data parsed from Excel file.
      if ($i == 0) {
			$subject=$elt;   // The title line
      } else {
		  
				$invoice="";
					$invoice=$invoice.'<?xml version="1.0" encoding="UTF-8"?>';
					$invoice=$invoice.$enterkey;
					$invoice=$invoice.'<invoice link="string">';
					$invoice=$invoice.$enterkey;
				$invoiceline="";
					$invoiceline=$invoiceline.'<?xml version="1.0" encoding="UTF-8"?>';
					$invoiceline=$invoiceline.$enterkey;
					$invoiceline=$invoiceline.'<invoice_line link="string">';
					$invoiceline=$invoiceline.$enterkey;
				$invoiceline2="";
					$invoiceline2=$invoiceline2.'<?xml version="1.0" encoding="UTF-8"?>';
					$invoiceline2=$invoiceline2.$enterkey;
					$invoiceline2=$invoiceline2.'<invoice_line link="string">';
					$invoiceline2=$invoiceline2.$enterkey;

		  $totalcount=count($elt);
		  $initial=0;
		  
		  for( $initial=1; $initial<$totalcount;$initial++){

			  switch ($initial){  //each case is to process the data from each colume in Excel file.
				  case 3:
				  case 6:
				  case 7:
				  case 9:
						$invoice=$invoice.modecomp($subject,$elt,$initial,$enterkey, $tabkey, 0,1);
						$invoice=$invoice.$enterkey;
						break;
						
				  case 10:
				     $invoice=$invoice.$tabkey."<additional_charges>";
				 	 $invoice=$invoice.$enterkey.str_repeat($tabkey,2)."<".$subject[$initial].">".$elt[$initial]. "</".$subject[$initial].">";
						for ($i=1; $i<5; $i++)
						{
							$invoice=$invoice.modesimple($subject,$elt,$initial,$i,$enterkey, $tabkey, 1,2);
						} 

					 $invoice=$invoice.$enterkey.$tabkey."</additional_charges>".$enterkey;
				    	$initial=$initial+4;
					break; 
					
                 case 15:
					 $invoice=$invoice.$tabkey."<invoice_vat>";
						   
						   for($i=0;$i<8;$i++){
							   
							   if($i==2 ||$i==5){
								   $invoice=$invoice.modecomp($subject,$elt,$initial+$i,$enterkey, $tabkey, 1,2);
							   }else{
								   $invoice=$invoice.modesimple($subject,$elt,$initial,$i, $enterkey, $tabkey, 1,2);
							   }
						   }
					 $invoice=$invoice.$enterkey.$tabkey."</invoice_vat>".$enterkey;
						$initial=$initial+7;
					break;	

                case 23:
                     $invoice=$invoice.$tabkey."<explicit_ratios>";
						 $invoice=$invoice.$enterkey.str_repeat($tabkey,2)."<explicit_ratio>";
						 	for($i=0;$i<2;$i++){
							   
							   if($i==0){
								   $invoice=$invoice.modecomp($subject,$elt,$initial+$i,$enterkey, $tabkey, 1,3);
							   }else{
								   $invoice=$invoice.modesimple($subject,$elt,$initial,$i, $enterkey, $tabkey, 1,3);
							   }
								
						   }
						 
						  $invoice=$invoice.$enterkey.str_repeat($tabkey,2)."</explicit_ratio>";
					  $invoice=$invoice.$enterkey.$tabkey."</explicit_ratios>".$enterkey;

					  $initial=$initial+1;
					 break;
					 
				case 25:
				       $invoice=$invoice.$tabkey."<payment>";
						 	for($i=0;$i<8;$i++){
							   
							   if($i==3 ||$i==7){
								  // echo "\r\n\t";
								  $invoice=$invoice.modecomp($subject,$elt,$initial+$i,$enterkey, $tabkey, 1,2);
							   }else{
								   
								  // echo "\r\n\t\t";
								   $invoice=$invoice.modesimple($subject,$elt,$initial,$i, $enterkey, $tabkey, 1,2);
							   }
								
						   }
					  $invoice=$invoice.$enterkey.$tabkey."</payment>".$enterkey;;

					  $initial=$initial+7;
					 break;
					 
				case 33:
				       $invoice=$invoice.$tabkey."<note>";
						 	for($i=0;$i<4;$i++){

								   $invoice=$invoice.modesimple($subject,$elt,$initial,$i, $enterkey, $tabkey, 1,2);
							
						   }
					  $invoice=$invoice.$enterkey.$tabkey."</note>".$enterkey;
					  $initial=$initial+3;
					 break;

                 case 37:
				 case 50:
				 case 51:
				 case 52:
                      $invoiceline=$invoiceline.modecomp($subject,$elt,$initial,$enterkey, $tabkey, 0,1);
						$invoiceline=$invoiceline.$enterkey;
						 break;
						
				case 38:
				case 39:
				case 40:
				case 41:
				case 42:
				case 43:				
				case 44:
				case 45:
				case 46:   
				case 47:
				case 48:
				case 49:
				case 53:
					$invoiceline=$invoiceline.$tabkey."<".$subject[$initial].">".$elt[$initial]. "</".$subject[$initial].">";
				    $invoiceline=$invoiceline.$enterkey;
					break;
					
				case 54:
					$invoiceline=$invoiceline.$tabkey."<invoice_line_vat>";
						 	for($i=0;$i<3;$i++){
							   
							   if($i==0){
								  // echo "\r\n\t";
								  $invoiceline=$invoiceline.modecomp($subject,$elt,$initial+$i,$enterkey, $tabkey, 1,2);
							   }else{
								   
								  // echo "\r\n\t\t";
								   $invoiceline=$invoiceline.modesimple($subject,$elt,$initial,$i, $enterkey, $tabkey, 1,2);
							   }
								
						   }
					  $invoiceline=$invoiceline.$enterkey.$tabkey."</invoice_line_vat>".$enterkey;;

					  $initial=$initial+2;
					 break;
					 
				case 57:
                     $invoiceline=$invoiceline.$tabkey."<fund_distributions>";
						 $invoiceline=$invoiceline.$enterkey.str_repeat($tabkey,2)."<fund_distribution>";
						 	for($i=0;$i<3;$i++){
							   
							   if($i==0){
								   $invoiceline=$invoiceline.modecomp($subject,$elt,$initial+$i,$enterkey, $tabkey, 1,3);
							   }else{
								   $invoiceline=$invoiceline.modesimple($subject,$elt,$initial,$i, $enterkey, $tabkey, 1,3);
							   }
						   }
						 
						$invoiceline=$invoiceline.$enterkey.str_repeat($tabkey,2)."</fund_distribution>";
					    $invoiceline=$invoiceline.$enterkey.$tabkey."</fund_distributions>".$enterkey;

						$initial=$initial+2;
					 break;
					 
					 
				 case 60:
				 case 73;
				 case 74:
				 case 75:
                      $invoiceline2=$invoiceline2.modecomp($subject,$elt,$initial,$enterkey, $tabkey, 0,1);
						$invoiceline2=$invoiceline2.$enterkey;
						 break;
						
				case 61:
				case 62:
				case 63:
				case 64:
				case 65:
				case 66:				
				case 67:
				case 68:
				case 69:   
				case 70:
				case 71:
				case 72:
				case 76:
					$invoiceline2=$invoiceline2.$tabkey."<".$subject[$initial].">".$elt[$initial]. "</".$subject[$initial].">";
				    $invoiceline2=$invoiceline2.$enterkey;
					break;
					
				case 77:
					$invoiceline2=$invoiceline2.$tabkey."<invoice_line_vat>";
						 	for($i=0;$i<3;$i++){
							   
							   if($i==0){
								  // echo "\r\n\t";
								  $invoiceline2=$invoiceline2.modecomp($subject,$elt,$initial+$i,$enterkey, $tabkey, 1,2);
							   }else{
								   
								  // echo "\r\n\t\t";
								   $invoiceline2=$invoiceline2.modesimple($subject,$elt,$initial,$i, $enterkey, $tabkey, 1,2);
							   }
								
						   }
					  $invoiceline2=$invoiceline2.$enterkey.$tabkey."</invoice_line_vat>".$enterkey;;

					  $initial=$initial+2;
					 break;
					 
				case 80:
                     $invoiceline2=$invoiceline2.$tabkey."<fund_distributions>";
						 $invoiceline2=$invoiceline2.$enterkey.str_repeat($tabkey,2)."<fund_distribution>";
						 	for($i=0;$i<3;$i++){
							   
							   if($i==0){
								   $invoiceline2=$invoiceline2.modecomp($subject,$elt,$initial+$i,$enterkey, $tabkey, 1,3);
							   }else{
								   $invoiceline2=$invoiceline2.modesimple($subject,$elt,$initial,$i, $enterkey, $tabkey, 1,3);
							   }
								
						   }
						 
						$invoiceline2=$invoiceline2.$enterkey.str_repeat($tabkey,2)."</fund_distribution>";
					    $invoiceline2=$invoiceline2.$enterkey.$tabkey."</fund_distributions>".$enterkey;

						$initial=$initial+2;
					 break;				
			
				
				default:
				   $invoice=$invoice.$tabkey."<".$subject[$initial].">".$elt[$initial]. "</".$subject[$initial].">";
				  $invoice=$invoice.$enterkey;
			   }
		  }
		  				$invoice=$invoice."</invoice>";
						$invoiceline=$invoiceline."</invoice_line>";
						$invoiceline2=$invoiceline2."</invoice_line>";
						echo $invoice;
						echo $enterkey;
						echo $invoiceline;
						echo $enterkey;
						echo $invoiceline2;
						
						echo $enterkey;
						$invoiceline1_xml = simplexml_load_string($invoiceline);
						$type=$invoiceline1_xml->type->xml_value;
						echo $type;
						echo $enterkey;
						$invoiceline2_xml = simplexml_load_string($invoiceline2);
						$type=$invoiceline2_xml->type->xml_value;
						echo $type;
						
						$invoice_id=apicall_invoice($invoice,$enterkey,$tabkey);
						echo $enterkey;
						Echo $invoice_id;
						
						apicall_invoiceline($invoiceline,$invoice_id,$enterkey,$tabkey);
						apicall_invoiceline($invoiceline2,$invoice_id,$enterkey,$tabkey);
						apicall_process_invoice($invoice_id,$enterkey,$tabkey);	  
     

      }    
	   $i++;

    }

  } else {
   $invoice=$invoice.SimpleXLSX::parseError();
  }



function apicall_invoice($invoice_para,$enterkey, $tabkey){  // The API used to create invoice
		$ch = curl_init();
		$baseUrl = 'https://api-na.hosted.exlibrisgroup.com/almaws/v1/acq/invoices';
		$queryParams = array(
			'apikey' => '' // API key. Please put your institution's API key here. If you use sandbox API key, it will put user records into sandbox. If you use production API key, it will put user records into production environment. Ex Libris automatically puts data into Sandbox or Production according to API key type.
		);
		$url = $baseUrl . "?" . http_build_query($queryParams);
		echo "url: ". $url;
		curl_setopt($ch, CURLOPT_URL, $url);
		curl_setopt($ch, CURLOPT_RETURNTRANSFER, TRUE);
		curl_setopt($ch, CURLOPT_HEADER, TRUE);
		curl_setopt($ch, CURLOPT_CUSTOMREQUEST, 'POST');  //We use PUT to update holding records
		curl_setopt($ch, CURLOPT_POSTFIELDS, $invoice_para);
		curl_setopt($ch, CURLOPT_HTTPHEADER, array('Content-Type: application/xml'));
		//curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, TRUE);
		curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, FALSE);
		
		echo "$ch:" . $ch;
			$response = curl_exec($ch);
		
		if (curl_errno($ch)) {
			// This would be your first hint that something went wrong
			die('Couldn\'t send request: ' . curl_error($ch));
		} else {
			// Check the HTTP status code of the request
			$resultStatus = curl_getinfo($ch, CURLINFO_HTTP_CODE);
			if ($resultStatus == 200) {
				echo "Everything went better than expected";
				// Everything went better than expected
			} else {
			die('Request failed: HTTP status code: ' . $resultStatus);
			}
		}
	//	echo "   Responce:    ".$response;
		
		$header_size = curl_getinfo($ch, CURLINFO_HEADER_SIZE);
		$header = substr($response, 0, $header_size);
		$body = substr($response, $header_size);
		
	//	echo "body: ".$body;
		
		$xml = simplexml_load_string($body);
		$id= $xml->id;
		echo $enterkey;
		echo $id;
		return $id;
		
		curl_close($ch);
	}
	
function apicall_invoiceline($invoiceline_para,$invoice_id, $enterkey, $tabkey){  // The API used to create invoice line.
		$ch = curl_init();
		$baseUrl = 'https://api-na.hosted.exlibrisgroup.com/almaws/v1/acq/invoices/';
		$queryParams = array(
			'apikey' => '' // API key. Please put your institution's API key here. If you use sandbox API key, it will put user records into sandbox. If you use production API key, it will put user records into production environment. Ex Libris automatically puts data into Sandbox or Production according to API key type.
		);
		$url = $baseUrl . $invoice_id."/lines?" . http_build_query($queryParams);
		echo $enterkey."url: ". $url;
		curl_setopt($ch, CURLOPT_URL, $url);
		curl_setopt($ch, CURLOPT_RETURNTRANSFER, TRUE);
		curl_setopt($ch, CURLOPT_HEADER, TRUE);
		curl_setopt($ch, CURLOPT_CUSTOMREQUEST, 'POST');  //We use PUT to update holding records
		curl_setopt($ch, CURLOPT_POSTFIELDS, $invoiceline_para);
		curl_setopt($ch, CURLOPT_HTTPHEADER, array('Content-Type: application/xml'));
		//curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, TRUE);
		curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, FALSE);
		
			echo "$ch:" . $ch;
			$response = curl_exec($ch);
			
			if (curl_errno($ch)) {
				// This would be your first hint that something went wrong
				die($enterkey.'Couldn\'t send request: ' . curl_error($ch));
			} else {
				// Check the HTTP status code of the request
				$resultStatus = curl_getinfo($ch, CURLINFO_HTTP_CODE);
				if ($resultStatus == 200) {
					echo $enterkey."Everything went better than expected";
					// Everything went better than expected
				} else {
				die($enterkey.'Request failed: HTTP status code: ' . $resultStatus);
				}
			}
		//	echo "   Responce:    ".$response;
		
		$header_size = curl_getinfo($ch, CURLINFO_HEADER_SIZE);
		$header = substr($response, 0, $header_size);
		$body = substr($response, $header_size);
		
		//	echo "body: ".$body;
		/*	
			$xml = simplexml_load_string($body);
			$id= $xml->id;
			echo $enterkey;
			echo $id;
			
			curl_close($ch);*/
	}
	
	
function apicall_process_invoice($invoiceid,$enterkey,$tabkey){ //The API used to process invoice
	$ch = curl_init();
	$baseUrl = 'https://api-na.hosted.exlibrisgroup.com/almaws/v1/acq/invoices/';
	$queryParams = array(
		'apikey' => '' // API key. Please put your institution's API key here. If you use sandbox API key, it will put user records into sandbox. If you use production API key, it will put user records into production environment. Ex Libris automatically puts data into Sandbox or Production according to API key type.
	);
	$url = $baseUrl . $invoiceid."?" . "op=process_invoice&". http_build_query($queryParams);
	curl_setopt($ch, CURLOPT_URL, $url);
		$response = curl_exec($ch);
	
		if (curl_errno($ch)) {
			// This would be your first hint that something went wrong
			die($enterkey.'Couldn\'t send request: ' . curl_error($ch));
		} else {
			// Check the HTTP status code of the request
			$resultStatus = curl_getinfo($ch, CURLINFO_HTTP_CODE);
			if ($resultStatus == 200) {
				echo $enterkey."Everything went better than expected.";
			} else {
			die($enterkey.'Request failed: HTTP status code: ' . $resultStatus);
			}
		}
	curl_close($ch);
    }
	
  
function  modesimple($sub, $elt, $ini, $ii, $enter, $tab, $enterN, $tabN) {    //The function used to generate XML tag and indent
		return str_repeat($enter,$enterN). str_repeat($tab, $tabN)."<".$sub[$ini+$ii].">".$elt[$ini+$ii]. "</".$sub[$ini+$ii].">";
	}

function  modecomp($sub, $elt, $ini, $enter, $tab, $enterN, $tabN) {   //The function used to generate complicated XML tag and indent
		//$enterkey="\r\n";  //local variable
		//$tabkey="\t";       //local variable
		  if ($enterN==0)
		  {
			  return str_repeat($tab, $tabN)."<".$sub[$ini].">".$enter.str_repeat($tab, $tabN+1)."<xml_value>".$elt[$ini]. "</xml_value>".$enter.str_repeat($tab, $tabN)."</".$sub[$ini].">";
		  }else{
			  return str_repeat($enter,$enterN).str_repeat($tab, $tabN)."<".$sub[$ini].">". str_repeat($enter,$enterN).str_repeat($tab, $tabN+1)."<xml_value>".$elt[$ini]. "</xml_value>".str_repeat($enter,$enterN).str_repeat($tab, $tabN)."</".$sub[$ini].">";
		  }
		
	}	

?>

