<?php

namespace App\Command;

use App\Services\AllBusiness;
use App\Services\Business;
use App\Services\ExtractFromHtml;
use App\Services\InsertInSpreadsheet;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Output\OutputInterface;
use Symfony\Component\Console\Style\SymfonyStyle;
use Symfony\Component\DependencyInjection\ParameterBag\ParameterBagInterface;

class BourseSpreadsheetCommand extends Command
{
    protected static $defaultName = 'app:bourse-spreadsheet';
    protected static $defaultDescription = 'Get spreadsheets from data';

    // Construction url
    protected static $yahooUrl = "https://fr.finance.yahoo.com/quote/";
    protected $whichData = [ "financials" => "/financials",
                             "balanceSheet" => "/balance-sheet",
                             "cashFlow" => "/cash-flow" ];
//    protected $stateAppendingYahooBinance = ["DOWJONES" => "", "AEX25" => "", "BEL20" => "", "ALLCAC" => ".PA", "CAC40"  => ".PA", "SBF120" => ".PA", "NASDAQ100" => "", "SANDB500" => ""];

    // Target data (array website)
    protected $yahooThClassname = "D(tbr)";
    protected $yahooTdClassName = "rw-expnded";

    private $params;

    public function __construct(ParameterBagInterface $params) {
        $this->params = $params;
        parent::__construct();
    }

    protected function configure()
    {
        $this
            ->setDescription(self::$defaultDescription)
        ;
    }

    protected function execute(InputInterface $input, OutputInterface $output): int
    {
        $io = new SymfonyStyle($input, $output);
//        $allBusiness = (new Business)->getAllBusiness();
        $count = 0;

//        foreach ($allBusiness as $category => $allBusinessByCategory){
             foreach (AllBusiness::$allBusiness as $business) {
                 $index = 0;
                foreach ($this->whichData as $extendUrl) {
                    // get content page
                    $page = @file_get_contents(self::$yahooUrl . $business . $extendUrl);
                    if ($page !== false) {
                        // setup document
                        $page = mb_convert_encoding($page, 'HTML-ENTITIES', 'UTF-8');
                        $documentHTML = new ExtractFromHtml($page);
                        $sheet = new InsertInSpreadsheet($this->params, $extendUrl);

                        $title = $documentHTML->getTitlePage();
                        $thTable = $documentHTML->getTitlesData($this->yahooThClassname);
                        $tdTable = $documentHTML->getData($this->yahooTdClassName);

                        if (!empty($title) && !empty($thTable) && !empty($tdTable)) {
                            $sheet->insertDataInSpreadsheet($title, $thTable, $tdTable);
                        } else {
                            $fileError = $this->params->get('kernel.project_dir')."/public/spreadsheets/errors.txt";
                            $current = @file_get_contents($fileError);
                            if ($current !== false) {
                                $current .= self::$yahooUrl . $business . $extendUrl . "\n";
                                file_put_contents($fileError, $current);
                            }
                        }
                        $count ++;
                        $index++;
                        echo $count."-";
                    } else {
                        $fileError = $this->params->get('kernel.project_dir')."/public/spreadsheets/errors.txt";
                        $current = @file_get_contents($fileError);
                        if ($current !== false) {
                            $current .= self::$yahooUrl . $business . $extendUrl . "\n";
                            file_put_contents($fileError, $current);
                        }
                    }
                 }
             }
//         }

        $io->success('Spreadsheets created');
        return Command::SUCCESS;
    }
}
