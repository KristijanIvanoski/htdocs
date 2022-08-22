<?php

$vars = $_GET;
if(empty($vars['mailSrv'])){
	echo "No mailSrv specified";
	exit;
}
if (empty($vars['mailUser'])) {
	echo "No mailUser specified";
	exit;
}
if (empty($vars['mailPass'])) {
	echo "No mailSrv specified";
	exit;
}
if (empty($vars['subjectFilter'])) {
	echo "No mailSrv specified";
	exit;
}

$mailbox = "{". $vars['mailSrv'] .":993/imap/ssl/novalidate-cert}INBOX";
$username = $vars['mailUser'];
$password = $vars['mailPass'];

$inbox = imap_open($mailbox, $username, $password) or die('Cannot connect to email: ' . imap_last_error());

$emails = imap_search($inbox, $vars['subjectFilter']);

if(empty($emails)){
	echo "No emails found";
	exit;
}
rsort($emails);

$attachmentsList = [];

function deleteDirectory($dir)
{
	if (!file_exists($dir)) {
		return true;
	}

	if (!is_dir($dir)) {
		return unlink($dir);
	}

	foreach (scandir($dir) as $item) {
		if ($item == '.' || $item == '..') {
			continue;
		}

		if (!deleteDirectory($dir . DIRECTORY_SEPARATOR . $item)) {
			return false;
		}
	}

	return rmdir($dir);
}

deleteDirectory('..\..\attachment');

/* if any emails found, iterate through each email */
if ($emails) {


	$count = 1;

	/* put the newest emails on top */
	rsort($emails);

	/* for every email... */
	foreach ($emails as $email_number) {

		/* get information specific to this email */
		$overview = imap_fetch_overview($inbox, $email_number, 0);

		$message = imap_fetchbody($inbox, $email_number, 2);

		/* get mail structure */
		$structure = imap_fetchstructure($inbox, $email_number);

		$attachments = array();

		/* if any attachments found... */
		if (isset($structure->parts) && count($structure->parts)) {
			for ($i = 0; $i < count($structure->parts); $i++) {
				$attachments[$i] = array(
					'is_attachment' => false,
					'filename' => '',
					'name' => '',
					'attachment' => ''
				);

				if ($structure->parts[$i]->ifdparameters) {
					foreach ($structure->parts[$i]->dparameters as $object) {
						if (strtolower($object->attribute) == 'filename') {
							$attachments[$i]['is_attachment'] = true;
							$attachments[$i]['filename'] = $object->value;
						}
					}
				}

				if ($structure->parts[$i]->ifparameters) {
					foreach ($structure->parts[$i]->parameters as $object) {
						if (strtolower($object->attribute) == 'name') {
							$attachments[$i]['is_attachment'] = true;
							$attachments[$i]['name'] = $object->value;
						}
					}
				}

				if ($attachments[$i]['is_attachment']) {
					$attachments[$i]['attachment'] = imap_fetchbody($inbox, $email_number, $i + 1);

					/* 3 = BASE64 encoding */
					if ($structure->parts[$i]->encoding == 3) {
						$attachments[$i]['attachment'] = base64_decode($attachments[$i]['attachment']);
					}
					/* 4 = QUOTED-PRINTABLE encoding */ elseif ($structure->parts[$i]->encoding == 4) {
						$attachments[$i]['attachment'] = quoted_printable_decode($attachments[$i]['attachment']);
					}
				}
			}
		}

		/* iterate through each attachment and save it */

		foreach ($attachments as $attachment) {
			if ($attachment['is_attachment'] == 1) {
				$filename = $attachment['name'];

				if(!str_contains($filename, '.xlsx')){
					continue;				}

				if (empty($filename)) $filename = $attachment['filename'];

				if (empty($filename)) $filename = time() . ".dat";
				$folder = "../../attachment";
				if (!is_dir($folder)) {
					mkdir($folder);
				}
				$FName = $email_number . "-x-" . $filename;
				$fp = fopen("./" . $folder . "/" . $FName, "w+");
				array_push($attachmentsList, $FName);
				fwrite($fp, $attachment['attachment']);
				fclose($fp);
			}
		}
	}
}

/* close the connection */
imap_close($inbox);

echo json_encode($attachmentsList) . "\n";




