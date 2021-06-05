<?php
namespace App\Services;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Symfony\Component\DependencyInjection\ParameterBag\ParameterBagInterface;

class InsertInSpreadsheet {
    private static $path;
    private $sheet;
    private $writer;
    private $countRow = 1;

    private static $alaphabetNumeric = [
        "1" => "A",
        "2" => "B",
        "3" => "C",
        "4" => "D",
        "5" => "E",
        "6" => "F",
        "7" => "G",
        "8" => "H",
        "9" => "I",
        "10" => "J",
        "11" => "K",
        "12" => "L",
        "13" => "M",
        "14" => "N",
        "15" => "O",
        "16" => "P",
        "17" => "Q",
        "18" => "R",
        "19" => "S",
        "20" => "T",
        "21" => "U",
        "22" => "V",
        "23" => "W",
        "24" => "X",
        "25" => "Y",
        "26" => "Z"
    ];

    public function __construct(ParameterBagInterface $params, $keyExtend)
    {
        self::$path = $params->get('kernel.project_dir')."/public/spreadsheets$keyExtend-".date("Y-m-d").".xlsx";
        $this->checkIfFileExists();
    }

    private function checkIfFileExists() {
        if (file_exists(self::$path)) {
            $file = \PhpOffice\PhpSpreadsheet\IOFactory::load(self::$path);
            $this->sheet = $file->getActiveSheet();
            $this->writer = new Xlsx($file);
        } else {
            $spreadsheet= new Spreadsheet();
            $this->sheet = $spreadsheet->getActiveSheet();
            $this->writer = new Xlsx($spreadsheet);
        }
    }

    public function insertDataInSpreadsheet(string $title, array $thTable, array $tdTable): void {
        if (file_exists(self::$path)) {
            $this->countRow = $this->sheet->getHighestDataRow();
            $this->countRow += 2;
        }

        $this->sheet->setCellValue("A$this->countRow", $title);
        $this->countRow++;

        for ($i = 0; $i < count($thTable); $i++) {
            $this->sheet->setCellValue(self::$alaphabetNumeric[$i+1].$this->countRow, $thTable[$i]);
        }
        $this->countRow++;

        foreach ($tdTable as $value) {
            $tdTableReconstruct = $value;
            $countAlphabet = 1;
            for ($i = 0; $i < count($tdTableReconstruct); $i++) {
                if (!empty($tdTableReconstruct[$i]) && preg_match('/\d/', $tdTableReconstruct[$i]) !== 1 && $tdTableReconstruct[$i] !== "-" && $i !== 0) {
                    $this->countRow++;
                    $countAlphabet = 1;
                }
                $val =  preg_replace('/\s+/u', '', $tdTableReconstruct[$i]);

                if(is_numeric($val)) {
                    $this->sheet->setCellValue(self::$alaphabetNumeric[$countAlphabet] . $this->countRow, '.=VALUE'.(int)$val."'");
                } else {
                    $this->sheet->setCellValue(self::$alaphabetNumeric[$countAlphabet] . $this->countRow, $val);
                }

                $countAlphabet++;
            }
        }

        $this->writer->save(self::$path);
    }
}