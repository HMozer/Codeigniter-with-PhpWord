<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class TesteWord extends CI_Controller
{
	function index()
	{
		$this->load->library("Phpword");

		$phpWord = new \PhpOffice\PhpWord\PhpWord();
		$phpWord->getCompatibility()->setOoxmlVersion(14);
		$phpWord->getCompatibility()->setOoxmlVersion(15);

		$filename = 'test.docx';

		$section = $phpWord->addSection();
		$section->addText('TESTE 1', array('bold' => true,'underline' => 'single','name'=> 'arial','size' => 21,'color' =>'red'),array('align' => 'center', 'spaceAfter' => 10));

		$section->addTextBreak(1);
		$section->addTextBreak(1);
		$section->addText('FUNCINOU', array('name'=> 'arial','size' => 14),array('align' => 'left', 'spaceAfter' => 100));

		$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
		$objWriter->save($filename);


		// send results to browser to download
		header('Content-Description: File Transfer');
		header('Content-Type: application/octet-stream');
		header('Content-Disposition: attachment; filename='.$filename);
		header('Content-Transfer-Encoding: binary');
		header('Expires: 0');
		header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
		header('Pragma: public');
		header('Content-Length: ' . filesize($filename));
		flush();
		readfile($filename);
		unlink($filename); // deletes the temporary file
		exit;
	}
}