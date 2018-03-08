<?php

include('./doc.php');

class Docx_reader {

    public $savepath = 'docimages';
    public $doxImages;
    private $fileData = false;
    private $docPath = false;
    private $errors = array();
    private $styles = array();
    private $_rels = array();
    private $imagesArray = false;

    public function __construct() {
        // $this->doxImages = new DocxImages('doc.docx');
        $this->imgFolder = $t->savepath;
    }

    private function load($file) {
        if (file_exists($file)) {
            $zip = new ZipArchive();
            $openedZip = $zip->open($file);
            if ($openedZip === true) {
                //attempt to load styles:
                if (($styleIndex = $zip->locateName('word/styles.xml')) !== false) {


                    $stylesXml = $zip->getFromIndex($styleIndex);
                    $xml = simplexml_load_string($stylesXml);
                    $namespaces = $xml->getNamespaces(true);

                    $children = $xml->children($namespaces['w']);

                    foreach ($children->style as $s) {
                        $attr = $s->attributes('w', true);
                        if (isset($attr['styleId'])) {
                            $tags = array();
                            $attrs = array();
                            foreach (get_object_vars($s->rPr) as $tag => $style) {
                                $att = $style->attributes('w', true);
                                switch ($tag) {
                                    case "b":
                                        $tags[] = 'strong';
                                        break;
                                    case "i":
                                        $tags[] = 'em';
                                        break;
                                    case "color":
                                        //echo (String) $att['val'];
                                        $attrs[] = 'color:#' . $att['val'];
                                        break;
                                    case "sz":
                                        $attrs[] = 'font-size:' . ($att['val'] / 1.2) . 'px';
                                        break;
                                }
                            }
                            foreach (get_object_vars($s->pPr) as $tag => $style) {
                                $att = $style->attributes('w', true);
                                switch ($tag) {
                                    case "ind":
                                        $attrs['pPr'][] = 'padding-left:' . ($att['left'] / 1440) . 'in';
                                        break;
                                }
                            }
                            $styles[(String)$attr['styleId']] = array('tags' => $tags, 'attrs' => $attrs);
                        }
                    }
                    // print_r($styles);
                    $this->styles = $styles;
                }

                if (($indexRel = $zip->locateName('word/_rels/document.xml.rels')) !== false) {
                    // If found, read it to the string
                    $xml = $zip->getFromIndex($indexRel);
                    // Close archive file

                    $relXml = $zip->getFromIndex($indexRel);
                    $xml = simplexml_load_string($relXml);

                    $rels = $xml->children();

                    foreach($rels->Relationship as $r) {
                        $attrs = $r->attributes();
                        $target = (array)$attrs['Target'];
                        $_rels[(String)$attrs['Id']] = $target[0];
                    }
                    
                    $this->_rels = $_rels;

                    $this->saveFiles();
                    
                }

                if (($index = $zip->locateName('word/document.xml')) !== false) {
                    // If found, read it to the string
                    $data = $zip->getFromIndex($index);
                    // Close archive file
                    $zip->close();
                    return $data;
                }

                $zip->close();

            } else {
                switch($openedZip) {
                    case ZipArchive::ER_EXISTS:
                        $this->errors[] = 'File exists.';
                        break;
                    case ZipArchive::ER_INCONS:
                        $this->errors[] = 'Inconsistent zip file.';
                        break;
                    case ZipArchive::ER_MEMORY:
                        $this->errors[] = 'Malloc failure.';
                        break;
                    case ZipArchive::ER_NOENT:
                        $this->errors[] = 'No such file.';
                        break;
                    case ZipArchive::ER_NOZIP:
                        $this->errors[] = 'File is not a zip archive.';
                        break;
                    case ZipArchive::ER_OPEN:
                        $this->errors[] = 'Could not open file.';
                        break;
                    case ZipArchive::ER_READ:
                        $this->errors[] = 'Read error.';
                        break;
                    case ZipArchive::ER_SEEK:
                        $this->errors[] = 'Seek error.';
                        break;
                }
            }
        } else {
            $this->errors[] = 'File does not exist.';
        }
    }

    public function saveFiles() {
        $imagesSaver = new DocxImages($this->docPath);
        $this->imagesArray = $imagesSaver->saveAllImages($this->savepath);
    }

    public function setFile($arr) {
        $this->docPath = $arr['filepath'];
        $this->fileData = $this->load($arr['filepath']);
        $this->savepath = $arr["img_dir"];
    }

    public function to_plain_text() {
        if ($this->fileData) {
            return strip_tags($this->fileData);
        } else {
            return false;
        }
    }

    public function to_html() {
        if ($this->fileData) {

            // print_r(htmlspecialchars($this->fileData));
            
            // $fileIndex = array();
            
            // $files = $this->doxImages->indexes;

            // foreach($files as $k => $v) {
            //     $files[] = $k;
            //     unset($files[$k]);
            // }

            // foreach ($files as $value) {
            //     preg_match('/'.$value.'/', $this->fileData, $matches, PREG_OFFSET_CAPTURE);

            //     $fileIndex[] = $matches[0];
            // }
            
             // print_r($fileIndex);

            $xml = simplexml_load_string($this->fileData);

            $namespaces = $xml->getNamespaces(true);

            $children = $xml->children($namespaces['w']);

            // print_r($children);
            
            $html = '<!doctype html><html><head><meta http-equiv="Content-Type" content="text/html;charset=utf-8" /><title></title><style>span.block { display: block;  }</style></head><body>';

            foreach ($children->body->p as $p) {
                $style = '';
                
                $startTags = array();
                $startAttrs = array();
                
                if($p->pPr->pStyle) {                    
                    $objectAttrs = $p->pPr->pStyle->attributes('w',true);
                    $objectStyle = (String) $objectAttrs['val'];

                    if(isset($this->styles[$objectStyle])) {
                        $startTags = $this->styles[$objectStyle]['tags'];
                        $startAttrs = $this->styles[$objectStyle]['attrs'];
                    }

                    if(isset($this->styles[$objectStyle]['attrs']['pPr'])) {
                        foreach ($this->styles[$objectStyle]['attrs']['pPr'] as $attr) {
                            $style .= $attr . ';';
                        }
                    }
                }
                
                if ($p->pPr->spacing) {
                    $att = $p->pPr->spacing->attributes('w', true);
                    if (isset($att['before'])) {
                        $style.='padding-top:' . ($att['before'] / 10) . 'px;';
                    }
                    if (isset($att['after'])) {
                        $style.='padding-bottom:' . ($att['after'] / 10) . 'px;';
                    }
                }

                if ($p->pPr->jc) {
                    $att = $p->pPr->jc->attributes('w', true);
                    if (isset($att['val'])) {
                        $style.='text-align:' . $att['val'] . ';';
                    }
                }

                $html.='<span class="block" style="' . $style . '">';
                $li = false;
                if ($p->pPr->numPr) {
                    $li = true;
                    $html.='<li>';
                }


                   
                
                foreach ($p->r as $part) {
                    //echo $part->t;
                    $tags = $startTags;
                    $attrs = $startAttrs; 

                    if($part->pict && $this->imagesArray !== false) {
                        $v = $part->pict->children($namespaces['v']);

                        $imgStyle = $v->shape->attributes();

                        $imgRId = $v->shape->imagedata->attributes('r', true);

                        $keyPath = $this->imagesArray[(string)$this->_rels[(string)$imgRId->id]];

                        $imgSrc = './'. $this->savepath .'/' . 'media/' . $keyPath;

                        $html.= '<img src="'.$imgSrc.'" style="'.$imgStyle->style.'">';

                    }

                    if($part->drawing) {

                        $wp = $part->drawing->children($namespaces['wp']);

                        $imgAttrs = $wp->inline->extent->attributes();

                        $width = (int) $imgAttrs['cx'][0] * 96 / 914400;

                        $height = (int) $imgAttrs['cy'][0] * 96 / 914400;

                        $imgStyle = 'height:' . $height . 'px;width:' . $width . 'px;';

                        $aGraphic = $wp->inline->children($namespaces['a']);

                        $pic = $aGraphic->graphic->graphicData->children($namespaces['pic']);

                        $embed = $pic->pic->blipFill->children($namespaces['a']);

                        $resIdAttr = $embed->blip->attributes('r', true);

                        $keyPath = $this->imagesArray[(string)$this->_rels[(string)$resIdAttr->embed]];

                        $imgSrc = './'. $this->savepath .'/' . 'media/' . $keyPath;

                        $html.= '<img src="'.$imgSrc.'" style="'.$imgStyle.'">';


                    }

                    foreach (get_object_vars($part->pPr) as $k => $v) {
                        if ($k = 'numPr') {
                            $tags[] = 'li';
                        }
                    }

                    foreach (get_object_vars($part->rPr) as $tag => $style) {
                        //print_r($style->attributes());


                        $att = $style->attributes('w', true);
                        switch ($tag) {
                            case "b":
                                $tags[] = 'strong';
                                break;
                            case "i":
                                $tags[] = 'em';
                                break;
                            case "color":
                                //echo (String) $att['val'];
                                $attrs[] = 'color:#' . $att['val'];
                                break;
                            case "sz":
                                $attrs[] = 'font-size:' . ($att['val'] / 1.2) . 'px';
                                break;
                            case "rFonts":
                                $attrs[] = 'font-family:' . $att['val'];
                        }
                    }
                    $openTags = '';
                    $closeTags = '';
                    foreach ($tags as $tag) {
                        $openTags.='<' . $tag . '>';
                        $closeTags.='</' . $tag . '>';
                    }
                    $html.='<span style="' . implode(';', $attrs) . '">' . $openTags . $part->t . $closeTags . '</span>';
                }
                if ($li) {
                    $html.='</li>';
                }
                $html.="</span>";
            }

            //Trying to weed out non-utf8 stuff from the file:
            $regex = <<<'END'
/
  (
    (?: [\x00-\x7F]                 # single-byte sequences   0xxxxxxx
    |   [\xC0-\xDF][\x80-\xBF]      # double-byte sequences   110xxxxx 10xxxxxx
    |   [\xE0-\xEF][\x80-\xBF]{2}   # triple-byte sequences   1110xxxx 10xxxxxx * 2
    |   [\xF0-\xF7][\x80-\xBF]{3}   # quadruple-byte sequence 11110xxx 10xxxxxx * 3 
    ){1,100}                        # ...one or more times
  )
| .                                 # anything else
/x
END;
            preg_replace($regex, '$1', $html);

            return $html . '</body></html>';
            exit();
        }
    }

    public function get_errors() {
        return $this->errors;
    }

    private function getStyles() {
        
    }

}
