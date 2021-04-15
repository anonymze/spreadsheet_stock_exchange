<?php

namespace App\Services;

use DOMDocument;
use DOMXPath;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ExtractFromHtml {
    private $dom;
    private $finder;

    function __construct(string $page){
        $this->dom = new DOMDocument('1.0', 'utf-8');
        @$this->dom->loadHTML($page);
        $this->finder = new DomXPath($this->dom);
    }

    public function getTitlePage(): string
    {
            $title = $this->dom->getElementsByTagName('h1');
            return $title->item(0) !== null ? $title->item(0)->textContent : "";
    }

    public function getTitlesData(string $classname): array
    {
        $titles = [];
       if($this->finder->query("//*[contains(@class, '$classname')]")->length > 0) {
           $classNames = $this->finder->query("//*[contains(@class, '$classname')]")->item(0)->childNodes;

           foreach ($classNames as $textContent){
               $titles[] =  $textContent->textContent;
           }
       }
        return $titles;
    }

    public function getData(string $classname): array
    {
        $data = [];
        $singleTr = [];

        if($this->finder->query("//*[contains(@class, '$classname')]")->length > 0) {
            $trAll = $this->finder->query("//*[contains(@class, '$classname')]");

            foreach ($trAll as $column) {
                foreach ($column->childNodes->item(0)->childNodes as $textContent) {
                    $singleTr[] = $textContent->textContent;
                }
                $data[$singleTr[0]] = $singleTr;
            }
        }
        return $data;
    }
}