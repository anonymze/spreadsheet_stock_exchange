<?php
namespace App\Services;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
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

    protected $firstFormulaCell1 = "";
    protected $firstFormulaCell2 = "";
    protected $firstFormulaCell3 = "";
    protected $firstFormulaCell4 = "";
    protected $firstFormulaCell5 = "";
    protected $secondFormulaCell1 = "";
    protected $secondFormulaCell2 = "";
    protected $secondFormulaCell3 = "";
    protected $secondFormulaCell4 = "";
    protected $secondFormulaCell5 = "";

    protected $firstArrayFormulaCell = [];
    protected $secondArrayFormulaCell = [];

    protected $firstCountDown = 0;
    protected $secondCountDown = -1;

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

    public function insertDataInSpreadsheet(string $title, array $thTable, array $tdTable, int $count): void {
        if (file_exists(self::$path)) {
            $this->countRow = $this->sheet->getHighestDataRow();
            $this->countRow += 2;
        }

        $this->sheet->setCellValue("A$this->countRow", $title);
        $this->countRow++;

        for ($i = 0, $iMax = count($thTable); $i < $iMax; $i++) {
            $this->sheet->setCellValue(self::$alaphabetNumeric[$i+1].$this->countRow, $thTable[$i]);
        }
        $this->countRow++;

        foreach ($tdTable as $value) {
            $tdTableReconstruct = $value;
            $countAlphabet = 1;
            for ($i = 0, $iMax = count($tdTableReconstruct); $i < $iMax; $i++) {

                if (!empty($tdTableReconstruct[$i]) && preg_match('/\d/', $tdTableReconstruct[$i]) !== 1 && $tdTableReconstruct[$i] !== "-" && $i !== 0) {
                    $this->countRow++;
                    $countAlphabet = 1;
                }

                $val =  preg_replace('/-?\s?/u', '', $tdTableReconstruct[$i]);

                if ($this->firstCountDown > 0) {
                    $this->firstArrayFormulaCell[] = $val ?: "1";
                    $this->firstCountDown--;
                }

                if ($this->secondCountDown > 0) {
                    $this->secondArrayFormulaCell[] = $val ?: "1";
                    $this->secondCountDown--;
                }

                if ($this->secondCountDown === 0) {
                    $this->sheet->setCellValue(self::$alaphabetNumeric[$countAlphabet + 3] . ($this->countRow), "MARGE BRUTE");

                    for($c = 0, $cMax = count($this->firstArrayFormulaCell); $c < $cMax; $c++) {
                        $val1 = (int)$this->firstArrayFormulaCell[$c];
                        $val2 = (int)$this->secondArrayFormulaCell[$c];

                        if($val1 > 0 && $val2 > 0) {
                            $result = ($val2 / $val1) * 100;
                            if ($result > 65 && $result < 100) {
                                $this->sheet->setCellValue(self::$alaphabetNumeric[$countAlphabet + (4 + $c)] . $this->countRow, $result)
                                    ->getStyle(self::$alaphabetNumeric[$countAlphabet + (4 + $c)] . $this->countRow)
                                    ->getFill()
                                    ->setFillType(Fill::FILL_SOLID)
                                    ->getStartColor()
                                    ->setARGB('39a6a3');
                            } else {
                                $this->sheet->setCellValue(self::$alaphabetNumeric[$countAlphabet + (4 + $c)] . $this->countRow, $result);
                            }
                        }

                        if ($c === ($count - 2)) {
                            $this->secondCountDown = -1;
                        }
                    }


                }

                if ($val === "Chiffred’affairestotal") {
                    $this->setupFirstRows($countAlphabet, $this->countRow);
                    $this->firstCountDown = $count -1;
                    var_dump($this->firstCountDown);
                }

                if ($val === "Bénéficebrut") {
                    $this->setupSecondRows($countAlphabet, $this->countRow);
                    $this->secondCountDown = $count -1;
                }

                if(is_numeric($val)) {
                    $val = (int)$val;
                    $this->sheet->setCellValue(self::$alaphabetNumeric[$countAlphabet] . $this->countRow, "=VALUE($val)");
                } else {
                    $this->sheet->setCellValue(self::$alaphabetNumeric[$countAlphabet] . $this->countRow, $val);
                }

                $countAlphabet++;
            }
        }

        $this->writer->save(self::$path);
    }

    public function setupFirstRows($countAlphabet, $countRow) {
        $this->firstFormulaCell1 = self::$alaphabetNumeric[$countAlphabet + 1] . $countRow;
        $this->firstFormulaCell2 = self::$alaphabetNumeric[$countAlphabet + 2] . $countRow;
        $this->firstFormulaCell3 = self::$alaphabetNumeric[$countAlphabet + 3] . $countRow;
        $this->firstFormulaCell4 = self::$alaphabetNumeric[$countAlphabet + 4] . $countRow;
        $this->firstFormulaCell5 = self::$alaphabetNumeric[$countAlphabet + 5] . $countRow;
    }

    public function setupSecondRows($countAlphabet, $countRow) {
        $this->secondFormulaCell1 = self::$alaphabetNumeric[$countAlphabet + 1] . ($countRow);
        $this->secondFormulaCell2 = self::$alaphabetNumeric[$countAlphabet + 2] . ($countRow);
        $this->secondFormulaCell3 = self::$alaphabetNumeric[$countAlphabet + 3] . ($countRow);
        $this->secondFormulaCell4 = self::$alaphabetNumeric[$countAlphabet + 4] . ($countRow);
        $this->secondFormulaCell5 = self::$alaphabetNumeric[$countAlphabet + 5] . ($countRow);
    }

    public function setupThirdRows() {

    }
}