<html>
    <head>
        <meta charset="UTF-8">
    </head>
    <body>     
        <?php
            function exceltocsv_maint($inputFileName)
            {
                /** Set default timezone (will throw a notice otherwise) */
                date_default_timezone_set('Europe/Paris');
                include 'Classes/PHPExcel/IOFactory.php';

                //  Read your Excel workbook
                try {
                    $inputFileType = PHPExcel_IOFactory::identify($inputFileName);
                    $objReader = PHPExcel_IOFactory::createReader($inputFileType);
                    $objPHPExcel = $objReader->load($inputFileName);
                } catch (Exception $e) {
                    die('Error loading file "' . pathinfo($inputFileName, PATHINFO_BASENAME) 
                    . '": ' . $e->getMessage());
                }

                $sheet = $objPHPExcel->getSheet(0);
                $highestRow = $sheet->getHighestRow();
                $highestColumn = $sheet->getHighestColumn();
                $row = 10;
                
                $rowDataB = $sheet->rangeToArray('B' . $row . ':' . 'B' . $highestRow, NULL, TRUE, FALSE); 
                $rowDataF = $sheet->rangeToArray('F' . $row . ':' . 'F' . $highestRow, NULL, TRUE, FALSE); 
                $rowDataH = $sheet->rangeToArray('H' . $row . ':' . 'H' . $highestRow, NULL, TRUE, FALSE);
                $rowDataM = $sheet->rangeToArray('M' . $row . ':' . 'M' . $highestRow, NULL, TRUE, FALSE);   
                
                $fp = fopen('file.csv', 'w');
                foreach($rowDataF as $k=>$v)
                {
                    if ($rowDataB[$k][0] <> "N")
                    {
                        //echo $rowDataB[$k][0] . " | " . $v[0] . " | " . $rowDataM[$k][0] . " | " . $rowDataH[$k][0] . "<br />";
                        $rowDataFinal = array($v[0], $rowDataH[$k][0], $rowDataM[$k][0]);
                        
                        fputcsv($fp, $rowDataFinal, ";", "\"");
                    }
                }
                fclose($fp);

                //return $rowDataFinal;
            }
            
            $timestamp_debut = microtime(true);

            exceltocsv_maint('essai.xlsx');
            echo "Le fichier CSV (file.csv) vient d'être créé...";
            
            $timestamp_fin = microtime(true);
            $difference_ms = $timestamp_fin - $timestamp_debut;
            // affichage du résultat
            echo "<br /><br />Duree d'éxecution du script : " . $difference_ms . " secondes.";
            
        ?>
    </body>
</html>
