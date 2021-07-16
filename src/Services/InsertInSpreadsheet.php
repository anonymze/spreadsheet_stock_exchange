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

    protected $totalBusiness = "";
    protected $totalBusiness2 = "";
    protected $totalBusiness3 = "";
    protected $totalBusiness4 = "";
    protected $totalBusiness5 = "";
    protected $totalBusiness6 = "";
    protected $totalBusiness7 = "";
    protected $brutBusiness = "";
    protected $brutBusiness2 = "";
    protected $brutBusiness3 = "";
    protected $brutBusiness4 = "";
    protected $brutBusiness5 = "";
    protected $brutBusiness6 = "";
    protected $brutBusiness7 = "";
    protected $generalSellAdministrative = "";
    protected $generalSellAdministrative2 = "";
    protected $generalSellAdministrative3 = "";
    protected $generalSellAdministrative4 = "";
    protected $generalSellAdministrative5 = "";
    protected $generalSellAdministrative6 = "";
    protected $generalSellAdministrative7 = "";
    protected $openExploitation = "";
    protected $openExploitation1 = "";
    protected $openExploitation2 = "";
    protected $openExploitation3 = "";
    protected $openExploitation4 = "";
    protected $openExploitation5 = "";
    protected $openExploitation6 = "";
    protected $interestCharge = "";
    protected $interestCharge2 = "";
    protected $interestCharge3 = "";
    protected $interestCharge4 = "";
    protected $interestCharge5 = "";
    protected $interestCharge6 = "";
    protected $interestCharge7 = "";
    protected $netBenefice = "";
    protected $netBenefice2 = "";
    protected $netBenefice3 = "";
    protected $netBenefice4 = "";
    protected $netBenefice5 = "";
    protected $netBenefice6 = "";
    protected $netBenefice7 = "";
    protected $researchDevelopment = "";
    protected $researchDevelopment2 = "";
    protected $researchDevelopment3 = "";
    protected $researchDevelopment4 = "";
    protected $researchDevelopment5 = "";
    protected $researchDevelopment6 = "";
    protected $researchDevelopment7 = "";
    protected $ordinaryActionaryBenefice = "";
    protected $ordinaryActionaryBenefice2 = "";
    protected $ordinaryActionaryBenefice3 = "";
    protected $ordinaryActionaryBenefice4 = "";
    protected $ordinaryActionaryBenefice5 = "";
    protected $ordinaryActionaryBenefice6 = "";
    protected $ordinaryActionaryBenefice7 = "";
    protected $investFinance = "";
    protected $investFinance2 = "";
    protected $investFinance3 = "";
    protected $investFinance4 = "";
    protected $investFinance5 = "";
    protected $investFinance6 = "";
    protected $investFinance7 = "";
    protected $totalPassif  = "";
    protected $totalPassif2  = "";
    protected $totalPassif3  = "";
    protected $totalPassif4  = "";
    protected $totalPassif5  = "";
    protected $totalPassif6  = "";
    protected $totalPassif7  = "";
    protected $properCapitalTotal = "";
    protected $properCapitalTotal2 = "";
    protected $properCapitalTotal3 = "";
    protected $properCapitalTotal4 = "";
    protected $properCapitalTotal5 = "";
    protected $properCapitalTotal6 = "";
    protected $properCapitalTotal7 = "";
    protected $inventory = "";
    protected $inventory2 = "";
    protected $inventory3 = "";
    protected $inventory4 = "";
    protected $inventory5 = "";
    protected $inventory6 = "";
    protected $inventory7 = "";
    protected $arrangedBenefice = "";
    protected $arrangedBenefice2 = "";
    protected $arrangedBenefice3 = "";
    protected $arrangedBenefice4 = "";
    protected $arrangedBenefice5 = "";
    protected $arrangedBenefice6 = "";
    protected $arrangedBenefice7 = "";
    protected $longDept = "";
    protected $longDept2 = "";
    protected $longDept3 = "";
    protected $longDept4 = "";
    protected $longDept5 = "";
    protected $longDept6 = "";
    protected $longDept7 = "";
    protected $shortActifTotal = "";
    protected $shortActifTotal2 = "";
    protected $shortActifTotal3 = "";
    protected $shortActifTotal4 = "";
    protected $shortActifTotal5 = "";
    protected $shortActifTotal6 = "";
    protected $shortActifTotal7 = "";
    protected $shortPassifTotal = "";
    protected $shortPassifTotal2 = "";
    protected $shortPassifTotal3 = "";
    protected $shortPassifTotal4 = "";
    protected $shortPassifTotal5 = "";
    protected $shortPassifTotal6 = "";
    protected $shortPassifTotal7 = "";

    protected $firstCurrent = "";
    protected $secondCurrent = "";
    protected $firstCountDown = -1;
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

        $this->sheet->setCellValue("A$this->countRow", $title)
            ->getStyle("A$this->countRow")
            ->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('D3D3D3');
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

                /**
                 * global formula
                 */
                if ($this->firstCountDown >= 0 && $this->firstCountDown <= $count) {
                    ++$this->firstCountDown;
                    $val = $val ?: "1";

                    if ($this->firstCountDown === 1) {
                        $this->{$this->firstCurrent} = $val;
                    } else {
                        $this->{$this->firstCurrent.$this->firstCountDown} = $val;
                    }
                }


                if ($this->secondCountDown >= 0 && $this->secondCountDown < $count) {
                    ++$this->secondCountDown;
                    $val = $val ?: "1";

                    if ($this->secondCountDown === 1) {
                        $this->{$this->secondCurrent} = $val;
                    } else {
                        $this->{$this->secondCurrent.$this->secondCountDown} = $val;
                    }
                }

                if ($this->secondCountDown === $count - 1) {
                    for($c = 1; $c < $count; $c++) {
                        $color = false;

                        // check previous result and calculate %
                        if ($this->secondCurrent === "ordinaryActionaryBenefice" || $this->secondCurrent === "inventory" || $this->secondCurrent === "arrangedBenefice") {

                            // check augmentation every year
                            switch ($this->secondCurrent) {
                                case "ordinaryActionaryBenefice":
                                    $percent = 10;
                                    break;
                                case "arrangedBenefice":
                                case "inventory":
                                    $percent = 1;
                                    break;
                            }

                                if ($c === 1) {
                                    if ($this->{$this->secondCurrent} > 0 && $this->{$this->secondCurrent . ($c + 1)} > 0) {
                                        if ($this->secondCurrent === "longDept") {
                                            $calcul = $this->{$this->secondCurrent} < (4 * $this->{$this->firstCurrent});
                                        } else if ($this->secondCurrent === "shortPassifTotal") {
                                            $calcul = (1.5 * $this->{$this->firstCurrent}) > $this->{$this->secondCurrent};
                                        }
                                        else {
                                            $calcul = $this->{$this->secondCurrent} > ($this->{$this->secondCurrent . "2"} + ($this->{$this->secondCurrent . "2"} * $percent / 100));
                                        }

                                        if ($calcul) {
                                            $this->sheet->setCellValue(self::$alaphabetNumeric[$c + 1] . $this->countRow, $this->{$this->secondCurrent})
                                                ->getStyle(self::$alaphabetNumeric[$c + 1] . $this->countRow)
                                                ->getFill()
                                                ->setFillType(Fill::FILL_SOLID)
                                                ->getStartColor()
                                                ->setARGB('39a6a3');
                                        } else {
                                            $this->sheet->setCellValue(self::$alaphabetNumeric[$c + 1] . $this->countRow, $this->{$this->secondCurrent})
                                                ->getStyle(self::$alaphabetNumeric[$c + 1] . $this->countRow)
                                                ->getFill()
                                                ->setFillType(Fill::FILL_SOLID)
                                                ->getStartColor()
                                                ->setARGB('FF0000');
                                        }
                                    }
                                } else {
                                    if ($this->{$this->secondCurrent . $c} > 0 && $this->{$this->secondCurrent . ($c + 1)} > 0) {
                                        if ($this->secondCurrent === "longDept") {
                                            $calcul = $this->{$this->secondCurrent . $c} < (4 * $this->{$this->firstCurrent . $c});
                                        } else if ($this->secondCurrent === "shortPassifTotal") {
                                            $calcul = (1.5 * $this->{$this->firstCurrent . $c}) > $this->{$this->secondCurrent . $c};
                                        }
                                        else {
                                            $calcul = $this->{$this->secondCurrent . $c} > ($this->{$this->secondCurrent . ($c + 1)} + ($this->{$this->secondCurrent . ($c + 1)} * $percent / 100));
                                       }

                                        if ($calcul) {
                                            $this->sheet->setCellValue(self::$alaphabetNumeric[$c + 1] . $this->countRow, $this->{$this->secondCurrent . $c})
                                                ->getStyle(self::$alaphabetNumeric[$c + 1] . $this->countRow)
                                                ->getFill()
                                                ->setFillType(Fill::FILL_SOLID)
                                                ->getStartColor()
                                                ->setARGB('39a6a3');
                                        } else {
                                            $this->sheet->setCellValue(self::$alaphabetNumeric[$c + 1] . $this->countRow, $this->{$this->secondCurrent . $c})
                                                ->getStyle(self::$alaphabetNumeric[$c + 1] . $this->countRow)
                                                ->getFill()
                                                ->setFillType(Fill::FILL_SOLID)
                                                ->getStartColor()
                                                ->setARGB('FF0000');
                                        }
                                    }

                            }
                        } else {

                            // calcul formula basic
                            if ($c === 1) {
                                switch ($this->secondCurrent) {
                                    case "brutBusiness":
                                        $this->sheet->setCellValue(self::$alaphabetNumeric[$count + 3] . ($this->countRow), "MARGE BRUTE");
                                        break;
                                    case "generalSellAdministrative":
                                        $this->sheet->setCellValue(self::$alaphabetNumeric[$count + 3] . ($this->countRow), "FRAIS D'EXPLOIT");
                                        break;
                                    case "interestCharge":
                                        $this->sheet->setCellValue(self::$alaphabetNumeric[$count + 3] . ($this->countRow), "CHARGE D'INTERET");
                                        break;
                                    case "netBenefice":
                                        $this->sheet->setCellValue(self::$alaphabetNumeric[$count + 3] . ($this->countRow), "MARGE NET");
                                        break;
                                    case "researchDevelopment":
                                        $this->sheet->setCellValue(self::$alaphabetNumeric[$count + 3] . ($this->countRow), "RESERCHE DEVELOPPEMENT");
                                        break;
                                    case "investFinance":
                                        $this->sheet->setCellValue(self::$alaphabetNumeric[$count + 3] . ($this->countRow), "CAPEX");
                                        break;
                                    case "properCapitalTotal":
                                    $this->sheet->setCellValue(self::$alaphabetNumeric[$count + 3] . ($this->countRow), "RATIO FINANCIER");
                                    break;
                                }

                                if ($this->{$this->secondCurrent} > 0 && $this->{$this->firstCurrent} > 0) {
                                    if ($this->secondCurrent === "properCapitalTotal") {
                                       $result = ($this->{$this->secondCurrent} / $this->{$this->firstCurrent});
                                    } else {
                                        $result = ($this->{$this->secondCurrent} / $this->{$this->firstCurrent}) * 100;
                                    }
                                }
                            } else if ($this->{$this->secondCurrent . $c} > 0 && $this->{$this->firstCurrent . $c} > 0) {
                                if ($this->secondCurrent === "properCapitalTotal") {
                                    $result = ($this->{$this->secondCurrent . $c} / $this->{$this->firstCurrent . $c});
                                } else {
                                    $result = ($this->{$this->secondCurrent . $c} / $this->{$this->firstCurrent . $c}) * 100;
                                }
                            }

                            if (isset($result)) {
                                switch ($this->secondCurrent) {
                                    case "brutBusiness":
                                        $color = $result > 35;
                                        break;
                                    case "generalSellAdministrative":
                                        $color = $result < 30;
                                        break;
                                    case "interestCharge":
                                        $color = $result < 15;
                                        break;
                                    case "netBenefice":
                                        $color = $result > 15;
                                        break;
                                    case "researchDevelopment":
                                        $color = $result < 5;
                                        break;
                                    case "investFinance":
                                        $color = ($result > -25) ? true : ($result < -25 && $result > -50) ? "orange" : "red";
                                        break;
                                    case "properCapitalTotal":
                                        $color = $result < 0.8;
                                         break;
                                }

                                if ($color === true) {
                                    $this->sheet->setCellValue(self::$alaphabetNumeric[$count + (3 + $c)] . $this->countRow, $result)
                                        ->getStyle(self::$alaphabetNumeric[$count + (3 + $c)] . $this->countRow)
                                        ->getFill()
                                        ->setFillType(Fill::FILL_SOLID)
                                        ->getStartColor()
                                        ->setARGB('39a6a3');
                                } else if ($color === "orange"){
                                    $this->sheet->setCellValue(self::$alaphabetNumeric[$count + (3 + $c)] . $this->countRow, $result)
                                        ->getStyle(self::$alaphabetNumeric[$count + (3 + $c)] . $this->countRow)
                                        ->getFill()
                                        ->setFillType(Fill::FILL_SOLID)
                                        ->getStartColor()
                                        ->setARGB('FF7F00');

                                } else if ($color === "red") {
                                    $this->sheet->setCellValue(self::$alaphabetNumeric[$count + (3 + $c)] . $this->countRow, $result)
                                        ->getStyle(self::$alaphabetNumeric[$count + (3 + $c)] . $this->countRow)
                                        ->getFill()
                                        ->setFillType(Fill::FILL_SOLID)
                                        ->getStartColor()
                                        ->setARGB('FF0000');
                                }
                                else {
                                    $this->sheet->setCellValue(self::$alaphabetNumeric[$countAlphabet + (3 + $c)] . $this->countRow, $result);
                                }
                            }
                        }
                    }

                    $this->secondCountDown = -1;
                    $this->firstCountDown = -1;
                }

                /**
                 * FIRST FORMULA MARGE BRUT
                 */
                if ($val === "Chiffred’affairestotal") {
                    $this->firstCountDown = 0;
                    $this->firstCurrent = "totalBusiness";
                }

                if ($val === "Bénéficebrut") {
                    $this->secondCountDown = 0;
                    $this->secondCurrent = "brutBusiness";
                }

                /**
                 * SECOND FORMULA FRAIS D'EXPLOITATION
                 */
                if ($val === "Ventesgénéralesetadministratives") {
                    $this->secondCountDown = 0;
                    $this->firstCurrent = "brutBusiness";
                    $this->secondCurrent = "generalSellAdministrative";
                }

                /**
                 * THIRD FORMULA FRAIS D'EXPLOITATION
                 */
                if ($val === "Bénéficeouperted’exploitation") {
                    $this->firstCountDown = 0;
                    $this->firstCurrent = "openExploitation";
                }

                if ($val === "Charged’intérêt") {
                    $this->secondCountDown = 0;
                    $this->secondCurrent = "interestCharge";
                }

                /**
                 * FOURTH FORMULA MARGE NET
                 */
                if ($val === "Bénéficenet") {
                    $this->secondCountDown = 0;
                    $this->firstCurrent = "totalBusiness";
                    $this->secondCurrent = "netBenefice";
                }

                /**
                 * FIFTH FORMULA R AND D
                 */
                if ($val === "Développementdelarecherche") {
                    $this->secondCountDown = 0;
                    $this->firstCurrent = "brutBusiness";
                    $this->secondCurrent = "researchDevelopment";
                }

                /**
                 * SIXTH FORMULA ACTIONARY ORDINARY NET BENEFICE
                 */
                if ($val === "Bénéficenetdisponiblepourlesactionnairesordinaires") {
                    $this->secondCountDown = 0;
                    $this->secondCurrent = "ordinaryActionaryBenefice";
                }

                /**
                 * SEVENTH FORMULA CAPEX
                 */
                if ($val === "Dépensesd’investissement") {
                    $this->secondCountDown = 0;
                    $this->firstCurrent = "netBenefice";
                    $this->secondCurrent = "investFinance";
                }



                /**
                 * NINTH arrangedBenefice
                 */
                if ($val === "Bénéficesnonrépartis") {
                    $this->secondCountDown = 0;
                    $this->secondCurrent = "arrangedBenefice";
                }

                /**
                 * TENTH arrangedBenefice
                 */
                if ($val === "Inventaire") {
                    $this->secondCountDown = 0;
                    $this->secondCurrent = "inventory";
                }

                /**
                 * ELEVENTH passif totaux / totaldescapitauxpropres (pas de *100)
                 */
                if ($val === "Passifstotaux") {
                    $this->firstCountDown = 0;
                    $this->firstCurrent = "totalPassif";
                }

                if ($val === "Totaldescapitauxpropres") {
                    $this->secondCountDown = 0;
                    $this->secondCurrent = "properCapitalTotal";
                }

                /**
                 * twelvth ROTC
                 */
//                if ($val === "Passifstotaux") {
//                    $this->firstCountDown = 0;
//                    $this->firstCurrent = "totalPassif";
//                }
//
//                if ($val === "Totaldescapitauxpropres") {
//                    $this->secondCountDown = 0;
//                    $this->secondCurrent = "properCapitalTotal";
//                }

                /**
                 * thirdth dette à long terme
                 */
                if ($val === "Detteàlongterme") {
                    $this->secondCountDown = 0;
                    $this->firstCurrent = "netBenefice";
                    $this->secondCurrent = "longDept";
                }

                /**
                 * fourteenth dette Total des actifs à court terme x 1,5 > Total des Passif à court terme
                 */
                if ($val === "Totaldesactifsàcourtterme") {
                    $this->firstCountDown = 0;
                    $this->firstCurrent = "shortActifTotal";
                }

                if ($val === "Totaldespassifsàcourtterme") {
                    $this->secondCountDown = 0;
                    $this->secondCurrent = "shortPassifTotal";
                }

                if(is_numeric($val)) {
                    $val = (int)$val;
                    $this->sheet->setCellValue(self::$alaphabetNumeric[$countAlphabet] . $this->countRow, "=VALUE($val)");

                // color specific squares
                } else if ($val === "Passif" || $val === "Actifs") {
                    $this->sheet->setCellValue(self::$alaphabetNumeric[$countAlphabet] . $this->countRow, $val)
                                    ->getStyle(self::$alaphabetNumeric[$countAlphabet] . $this->countRow)
                                    ->getFill()
                                    ->setFillType(Fill::FILL_SOLID)
                                    ->getStartColor()
                                    ->setARGB('87CEEB');
                } else {
                    $this->sheet->setCellValue(self::$alaphabetNumeric[$countAlphabet] . $this->countRow, $val);
                }

                $countAlphabet++;
            }
        }

        $this->writer->save(self::$path);
    }
}