<?php

// Author: PaLaddin

include('./docx_reader.php');

$doc_path = 'sample.docx';
$img_folder_path = 'docimages';

$doc = new Docx_reader();
$doc->setFile(array(
	"filepath" => $doc_path,
	"img_dir" => $img_folder_path
));

if(!$doc->get_errors()) {
    $html = $doc->to_html();
    $plain_text = $doc->to_plain_text();

    echo $html;
} else {
    echo implode(', ',$doc->get_errors());
}
echo "\n";

?>

<style>
	body {
		width: 1000px;
	    margin-left: auto;
	    margin-right: auto;
	    /*border: 1px solid black;*/
	    font-family: Calibri Light, serif;
	}
	span.block {
		min-height: 30px;
		margin: 15px 0;
	}

</style>