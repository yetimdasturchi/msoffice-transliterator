<?php

function translate($text, $rev=FALSE){
    $cyr  = array('а','б','в','г','д','е','ё','ж','з','и','й','к','л','м','н','о','п','р','с','т','у', 
            'ф','х','ц','ч','ш','щ','ъ', 'ы','ь', 'э', 'ю','я','А','Б','В','Г','Д','Е','Ж','З','И','Й','К','Л','М','Н','О','П','Р','С','Т','У',
            'Ф','Х','Ц','Ч','Ш','Щ','Ъ', 'Ы','Ь', 'Э', 'Ю','Я' );
    $lat = array( 'a','b','v','g','d','e','io','zh','z','i','y','k','l','m','n','o','p','r','s','t','u',
                'f' ,'h' ,'ts' ,'ch','sh' ,'sht' ,'a', 'i', 'y', 'e' ,'yu' ,'ya','A','B','V','G','D','E','Zh',
                'Z','I','Y','K','L','M','N','O','P','R','S','T','U',
                'F' ,'H' ,'Ts' ,'Ch','Sh' ,'Sht' ,'A' ,'Y' ,'Yu' ,'Ya' );

    return $rev ? str_replace($lat, $cyr, $text) : str_replace($cyr, $lat, $text);
}

function is_html( $string ) {
  return preg_match("/<[^<]+>/",$string,$m) != 0;
}

function transliterate_docx($filename, $alp='lat'){
    $zip = new ZipArchive;
    $fileToModify = 'word/document.xml';

    if ($zip->open($filename) === TRUE) {
        $oldContents = $zip->getFromName($fileToModify);
        
        $newContents = $oldContents;
        $newContents = preg_replace_callback('/<w:t.*?>([^<]+)<\/w:t>/', function ($row) use($alp){
            if (empty($row[1])) return $row[0];
            if(!is_html( $row[1])) {
                preg_match('/<w:t(.*?)>'.preg_quote($row[1], '/').'<\/w:t>/', $row[0], $submatch);
                $content = htmlspecialchars( ($alp == 'cyr') ? translate( html_entity_decode( $row[1] ), 1 ) : translate( html_entity_decode( $row[1] ) ), ENT_XML1, 'UTF-8');
                if (!empty($submatch[1])) {
                    $ret = '<w:t'.$submatch[1].'>'.$content.'</w:t>';
                }else{
                    $ret = '<w:t>'.$content.'</w:t>';
                }
                return $ret;
            }

            return $row[0];
        }, $newContents);

        $zip->deleteName($fileToModify);
        $zip->addFromString($fileToModify, $newContents);
        return $zip->close();
    }

    return FALSE;
}


function transliterate_xlsx($filename, $alp='lat') {
    $zip = new ZipArchive;
    if ($zip->open($filename) === TRUE) {
        $oldContents = $zip->getFromName('xl/sharedStrings.xml');
        $newContents = $oldContents;

        $newContents = preg_replace_callback('/<t.*?>([^<]+)<\/t>/', function ($row) use($alp){
            if (empty($row[1])) return $row[0];
            if(!is_html( $row[1])) {
                preg_match('/<t(.*)>(.*)<\/t>/', $row[0], $submatch);
                $content = htmlspecialchars( ($alp == 'cyr') ? translate( html_entity_decode( $row[1] ), 1 ) : translate( html_entity_decode( $row[1] ) ), ENT_XML1, 'UTF-8');
                if (!empty($submatch[1])) {
                    $ret = '<t'.$submatch[1].'>'.$content.'</t>';
                }else{
                    $ret = '<t>'.$content.'</t>';
                }
                return $ret;
            }

            return $row[0];
        }, $newContents);

        $zip->deleteName('xl/sharedStrings.xml');
        $zip->addFromString('xl/sharedStrings.xml', $newContents);

        $workbook = $zip->getFromName('xl/workbook.xml');
        $newWorkbook = $workbook;
        
        preg_match_all('/<sheet.*?name="(.*?)".*?\/>/m', $newWorkbook, $sheets, PREG_SET_ORDER);
        $sheets_count = count($sheets);
        if (!empty($sheets)) {
            $sheets_count = count($sheets);
            $t_sheet_contents = [];
            for ($i=1; $i <= $sheets_count; $i++) {
                $t_sheet_contents[] = [
                    'id' => $i,
                    'content' => $zip->getFromName('xl/worksheets/sheet'.$i.'.xml'),
                ];
            }
            foreach ($sheets as $sheet) {
                $content = htmlspecialchars( ($alp == 'cyr') ? translate( html_entity_decode( $sheet[1] ), 1 ) : translate( html_entity_decode( $sheet[1] ) ), ENT_XML1, 'UTF-8');
                $newWorkbook = str_replace($sheet[1], $content, $newWorkbook);
                foreach ($t_sheet_contents as $k => $t_sheet_content) {
                    $t_sheet_contents[$k]['content'] = str_replace($sheet[1], $content, $t_sheet_content['content']);
                }
                sleep(0.05);
            }
            foreach ($t_sheet_contents as $t_sheet_content) {
                $zip->deleteName('xl/worksheets/sheet'.$t_sheet_content['id'].'.xml');
                $zip->addFromString('xl/worksheets/sheet'.$t_sheet_content['id'].'.xml', $t_sheet_content['content']);
            }
        }
        $zip->deleteName('xl/workbook.xml');
        $zip->addFromString('xl/workbook.xml', $newWorkbook);
        return $zip->close();
    }

    return FALSE;
}

function transliterate_pptx($filename, $alp='lat') {
    $zip = new ZipArchive;
    if ($zip->open($filename) === TRUE) {
        $slide_number = 1; 
        while(($xml_index = $zip->locateName('ppt/slides/slide'.$slide_number.'.xml')) !== false){
            $slide_content = $zip->getFromIndex($xml_index);
            $slide_content = preg_replace_callback('/<a:t.*?>([^<]+)<\/a:t>/', function ($row) use($alp){
                if (empty($row[1])) return $row[0];
                
                if(!is_html( $row[1])) {
                    preg_match('/<a:t(.*)>(.*)<\/a:t>/', $row[0], $submatch);
                    
                    $content = htmlspecialchars( ($alp == 'cyr') ? translate( html_entity_decode( $row[1] ), 1 ) : translate( html_entity_decode( $row[1] ) ), ENT_XML1, 'UTF-8');
                    
                    if (!empty($submatch[1])) {
                        $ret = '<a:t'.$submatch[1].'>'.$content.'</a:t>';
                    }else{
                        $ret = '<a:t>'.$content.'</a:t>';
                    }
                    
                    return $ret; 
                }

                return $row[0];
            }, $slide_content);
            
            $zip->deleteName('ppt/slides/slide'.$slide_number.'.xml');
            $zip->addFromString('ppt/slides/slide'.$slide_number.'.xml', $slide_content);
            
            $slide_number++;
        }

        return $zip->close();
    }

    return FALSE;
}
