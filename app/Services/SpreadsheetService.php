<?php

namespace App\Services;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Illuminate\Support\Facades\Log;
use Illuminate\Support\Facades\Redis;
use PhpOffice\PhpSpreadsheet\Calculation\Calculation;
use App\Models\Design;
use Illuminate\Support\Collection;
use App\Models\InvoiceType;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

class SpreadsheetService
{
    public $temp = 0;

    private $cellMappings;
    private $fVariationArray;
    private $plitaVariationArray;
    private $metalAndPlasticVariationArray;
    private $exceptionalSheetsArray;

    public function __construct()
    {
        $this->cellMappings = [
            'fLenta' => [
                'D4' => 'lfLength',
                'D5' => 0.3,
                'D8' => 0.5,
                'D9' => 'lfAngleX',
                'D10' => 'lfAngleT',
                'D11' => 'lfAngleG',
                'D12' => 'lfAngle45',
                'D14' => 0.2,
                'D15' => 0.2,
                'D16' => 'mfSquare',
                'D18' => 'vfCount'
            ],
            'fVinta' => [
                'D44' => 'vfCount',
                'D45' => 'vfBalk',
                'D86' => 'outer_p',
                'D89' => 'lfAngleG'
            ],
            'brus' => [
                'C113' => 'baseD20RubF',
                'C114' => 'baseBalk1',
                'C115' => 'stolby'
            ],
            'brusFloors' => [
                'I112' => 'Sfl0',
                'I113' => 'Sfl1',
                'I114' => 'Sfl2',
                'I115' => 'Sfl3',
                'I116' => 'Sfl4'
            ],
            'ocb' => [
                'D169' => 'baseLength',
                'D170' => 'baseD20'
            ],
            'rSoft' => [
                'D283' => 'roofSquare',
                'D284' => 'srCherep',
                'D285' => 'srKover',
                'D286' => 'srKonK',
                'D287' => 'srMastika1',
                'D288' => 'srMastika',
                'D289' => 'mrKon',
                'D290' => 'srKonOneSkat',
                'D291' => 'srPlanVetr',
                'D292' => 'srPlanK',
                'D293' => 'srKapelnik',
                'D294' => 'mrEndv',
                'D295' => 'srGvozd',
                'D296' => 'mrSam70',
                'D297' => 'mrPack',
                'D298' => 'mrIzospanAM',
                'D299' => 'mrIzospanAM35',
                'D300' => 'mrLenta',
                'D301' => 'mrRokvul',
                'D302' => 'mrIzospanB',
                'D303' => 'mrIzospanB35',
                'D304' => 'mrPrimUgol',
                'D305' => 'mrPrimNakl',
                'D306' => 'srOSB',
                'D308' => 'srAero',
                'D309' => 'srAeroSkat',
                'D310' => 'stropValue'
            ],
            'pv' => [
                'D351' => 'pvPart1',
                'D352' => 'pvPart2',
                'D353' => 'pvPart3',
                'D354' => 'pvPart4',
                'D355' => 'pvPart5',
                'D356' => 'pvPart6',
                'D357' => 'pvPart7',
                'D358' => 'pvPart8',
                'D359' => 'pvPart9',
                'D360' => 'pvPart10',
                'D361' => 'pvPart11',
                'D362' => 'pvPart12',
                'D363' => 'pvPart13'
            ],
            'mv' => [
                'D367' => 'mvPart1',
                'D368' => 'mvPart2',
                'D369' => 'mvPart3',
                'D370' => 'mvPart4',
                'D371' => 'mvPart5',
                'D372' => 'mvPart6',
                'D373' => 'mvPart7',
                'D374' => 'mvPart8',
                'D375' => 'mvPart9',
                'D376' => 'mvPart10',
                'D377' => 'mvPart11',
                'D378' => 'mvPart12',
                'D379' => 'mvPart13'
            ],
            'srRoofSection' => [
                'L306' => 'srKonShir',
                'L307' => 'srKonOneSkat',
                'L311' => 'srEndn',
                'L312' => 'srEndv',
                'L313' => 'mrSam35',
                'L314' => 'srSam70',
                'L315' => 'srPack',
                'L316' => 'srIzospanAM',
                'L317' => 'srIzospanAM35',
                'L322' => 'srPrimUgol',
                'L323' => 'srPrimNakl'
            ]
        ];
        $this->fVariationArray = [
            ['600x300', 0.5, 0.3, 'fLenta600x300', 'fSVR600x300'],
            ['700x300', 0.6, 0.3, 'fLenta700x300', 'fSVR700x300'],
            ['800x300', 0.7, 0.3, 'fLenta800x300', 'fSVR800x300'],
            ['900x300', 0.8, 0.3, 'fLenta900x300', 'fSVR900x300'],
            ['1000x300', 0.9, 0.3, 'fLenta1000x300', 'fSVR1000x300'],
            ['600x400', 0.5, 0.4, 'fLenta600x400', 'fSVR600x400'],
            ['700x400', 0.6, 0.4, 'fLenta700x400', 'fSVR700x400'],
            ['800x400', 0.7, 0.4, 'fLenta800x400', 'fSVR800x400'],
            ['900x400', 0.8, 0.4, 'fLenta900x400', 'fSVR900x400'],
            ['1000x400', 0.9, 0.4, 'fLenta1000x400', 'fSVR1000x400'],
            ['600x500', 0.5, 0.5, 'fLenta600x500', 'fSVR600x500'],
            ['700x500', 0.6, 0.5, 'fLenta700x500', 'fSVR700x500'],
            ['800x500', 0.7, 0.5, 'fLenta800x500', 'fSVR800x500'],
            ['900x500', 0.8, 0.5, 'fLenta900x500', 'fSVR900x500'],
            ['1000x500', 0.9, 0.5, 'fLenta1000x500', 'fSVR1000x500'],
            ['700x600', 0.6, 0.6, 'fLenta700x600', 'fSVR700x600'],
            ['800x600', 0.7, 0.6, 'fLenta800x600', 'fSVR800x600'],
            ['900x600', 0.8, 0.6, 'fLenta900x600', 'fSVR900x600'],
            ['1000x600', 0.9, 0.6, 'fLenta1000x600', 'fSVR1000x600'],
        ];
        $this->plitaVariationArray = [
            ['0.2', 'fMono20'],
            ['0.25', 'fMono25'],
            ['0.3', 'fMono30'],
            ['0.35', 'fMono35'],
        ];
        $this->metalAndPlasticVariationArray = [
            ['Смета Мягкая', 'rSoftP', 'F92', 'K92'],
            ['Смета Мягкая', 'rSoftM', 'F77', 'K77'],
            ['Смета Железо', 'rMetalP', 'F111', 'K111'],
            ['Смета Железо', 'rMetalM', 'F96', 'K96'],
        ];
        $this->exceptionalSheetsArray = [
            'СВ-Рост', 'плита', 'лента', 'Железо', 'Мягкая'
        ];
    }
    private function checkIfSheetIsExceptional($sheetName) {
        foreach ($this->exceptionalSheetsArray as $exception) {
            if (strpos($sheetName, $exception) !== false) {
                return true;
            }
        }
        return false;
    }
    public function handle($filePath, $design=1, $multiple=false, $labour=true, $debug=1, $config=null) {
        try {
            $spreadsheet = IOFactory::createReader('Xlsx')->load($filePath);
        } catch (\Exception $e) {
            throw $e;
        }

        if ($config) {
            $newFilePath = $this->processConfiguredSheets($spreadsheet, $design, $config);
            return $newFilePath;
        }

        if ($multiple) {
            foreach ($design as $design) {
                $this->handlePriceIndexing($spreadsheet, $design);
            }
        }

        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        Calculation::getInstance($spreadsheet)->clearCalculationCache();
        $this->processDatasheet($spreadsheet, $design);
        $filename = $design->id . "_" . time();
        $newFilePath = storage_path('app/public/orders/' . $filename . '.xlsx');
        $writer->save($newFilePath);

        return $newFilePath;
    }

    private function getSheetsToCombine($config) {
        $sheetsToCombine = [];
        $extraCols = false;
        foreach ($config as $sheetName => $sheetConfig) {
            $sheet = InvoiceType::where('ref', $sheetConfig)->firstOrFail();
            $checkExtra = $this->checkExtraColumn($sheet->sheetname);
            if($checkExtra) {
                $extraCols = true;
            }
            if ($sheet->label != "rNone" && $sheet->label != "fNone") {
                $sheetsToCombine[$sheet->id] = [$sheet->label, $sheet->sheetname, $checkExtra, $sheet->title];
            }
        }
        $sheetsToCombine["extra"] = $extraCols;
        return $sheetsToCombine;
    }

    private function processConfiguredSheets($spreadsheet, $design, $config)
{
    Calculation::getInstance($spreadsheet)->clearCalculationCache();
    $this->processDatasheet($spreadsheet, $design);
    $newSpreadsheet = new Spreadsheet();
    $newSpreadsheet->removeSheetByIndex(0); // Remove default sheet
    $sheetsToCombine = $this->getSheetsToCombine($config);
    $extraCol = $sheetsToCombine["extra"];
    unset($sheetsToCombine["extra"]);
    $sheetIndex = 0;
    $newSheet = $newSpreadsheet->createSheet($sheetIndex);
    $newSheet->setTitle("Смета");
    $newSheetRow = 1;
    $lastIteration = false;
    $sumsSection = [];
    Calculation::getInstance($spreadsheet)->clearCalculationCache();
    foreach ($sheetsToCombine as $sheetItem) {
        foreach ($this->exceptionalSheetsArray as $exception) {
            if (strpos($sheetItem[1], $exception) !== false) {
                switch ($exception) {
                    case "СВ-Рост":
                        $sheet = $spreadsheet->getSheetByName("data");
                        foreach ($this->fVariationArray as $size) {
                            if ($sheetItem[3] == $size[0]) {
                                $sheet->setCellValue('D5', $size[2]);
                                $sheet->setCellValue('D8', $size[1]);
                            }
                        }
                        break;
                    case "плита":
                        $sheet = $spreadsheet->getSheetByName("data");
                        foreach ($this->plitaVariationArray as $size) {
                            if ($sheetItem[0] == $size[1]) {
                                $sheet->setCellValue('D87', $size[0]);
                            }
                        }
                        break;
                    case "лента":
                        $sheet = $spreadsheet->getSheetByName("data");
                        foreach ($this->fVariationArray as $size) {
                            if ($sheetItem[3] == $size[0]) {
                                $sheet->setCellValue('D5', $size[2]);
                                $sheet->setCellValue('D8', $size[1]);
                            }
                        }
                        break;
                    case "Железо":
                        $sheet = $spreadsheet->getSheetByName("Смета Железо");
                        $sheet->getStyle('G41')->getFont()->setSize(12);
                        //remove first character from C3 and add it back
                        $C3row = substr($sheet->getCell('C3')->getValue(), 2);
                        $C3col = $sheet->getCell('C3')->getValue();
                        $C3col = $C3col[1];
                        //dealing with Vodostok
                        if ($sheetItem[3] == "Железо, пластик") {
                            $startIndex = 97;
                            $endRow = 112;
                            $rowsToDelete = $endRow - $startIndex;
                            $sheet->removeRow($startIndex, $rowsToDelete);
                        } else if ($sheetItem[3] == "Железо, метал") {
                            $startIndex = 82;
                            $endRow = 97;
                            $rowsToDelete = $endRow - $startIndex;
                            $sheet->removeRow($startIndex, $rowsToDelete);
                        } else throw new Exception("Неизвестный тип водостока");
                        //removing rows for unnecessary metaList lines
                        $deletedRows = $rowsToDelete;
                        $metaList = ($design->metaList);
                        $size = count($metaList);
                        $startIndex = 42;
                        $endIndex = 63;
                        $rowsToDelete = $endIndex - $startIndex - $size +1;
                        $firstRowToDelete = $startIndex + $size;
                        $sheet->removeRow($firstRowToDelete, $rowsToDelete);
                        $deletedRows += $rowsToDelete;
                        $C3row -= $deletedRows;
                        $sheet->setCellValue('C3', "r" . $C3col . $C3row);
                        break;
                    case "Мягкая":
                        $sheet = $spreadsheet->getSheetByName("Смета Мягкая");
                        $C3row = substr($sheet->getCell('C3')->getValue(), 2);
                        $C3col = $sheet->getCell('C3')->getValue();
                        $C3col = $C3col[1];
                        //dealing with Vodostok
                        if ($sheetItem[3] == "Мягкая, пластик") {
                            $startIndex = 78;
                            $endRow = 93;
                            $rowsToDelete = $endRow - $startIndex + 1;
                            $sheet->removeRow($startIndex, $rowsToDelete);
                        } else if ($sheetItem[3] == "Мягкая, метал") {
                            $startIndex = 63;
                            $endRow = 78;
                            $rowsToDelete = $endRow - $startIndex;
                            $sheet->removeRow($startIndex, $rowsToDelete);
                        } else throw new Exception("Неизвестный тип водостока");
                        $C3row -= $rowsToDelete;
                        $sheet->setCellValue('C3', "r" . $C3col . $C3row);
                        break;
                }
            }
        }
    }
    Calculation::getInstance($spreadsheet)->clearCalculationCache();
    foreach ($sheetsToCombine as $sheet) {
        //is it the last iteration
        if ($sheetIndex == count($sheetsToCombine)-1) {
            $lastIteration = true;
        }
        $sheetTitle = $sheet[1];
        $skipCol = $sheet[2];
        $invoiceType = InvoiceType::where('sheetname', $sheet[1])->firstOrFail();
        $spec = $invoiceType->sheet_spec;
        $sheet = $spreadsheet->getSheetByName($sheetTitle);
        Log::info("Processing sheet: " . $sheetTitle . ". Skip col: " . ($skipCol+$extraCol));

        


        if ($spreadsheet->sheetNameExists($sheetTitle)) {
            if ($lastIteration) {
                //map all column widths for first M cols
                for ($i = "A"; $i <= "M"; $i++) {
                    $newSheet->getColumnDimension($i)->setWidth($sheet->getColumnDimension($i)->getWidth());
                }
            }
            $thisSheetWithExtraCols = false;
            if ((($skipCol+$extraCol) != 2) && isset($extraCol) && $extraCol == true) {
                //insert row after C3
                $sheet->insertNewColumnBefore('C', 1);
                $sheet->insertNewColumnBefore('I', 1);
                Log::info("Inserting row before C");
                $lastRow = $sheet->getCell('D3')->getValue();
                $thisSheetWithExtraCols = true;
            } else {
                Log::info("Not inserting row before C");
                $lastRow = $sheet->getCell('C3')->getValue();
            }
            $totalsBoxStart = substr($lastRow, 1);
            Log::info("Totals box start: " . $totalsBoxStart);
            $sumsSection = $this->updateRunningTotals($sheet, $totalsBoxStart, $sumsSection);
            if ($lastIteration) {
                Log::info("Last iteration");
                $lastRow = substr($lastRow, 2)+7;
            } else {
                Log::info("Not last iteration");
                $lastRow = substr($lastRow, 2)-1;
            }
            
            $firstNewSheetRow = $newSheetRow;
            $lastNewSheetRow = $newSheetRow + $lastRow;
            
            if ($sheetIndex == 0) {
                $row = 1;
            } else $row = 7;
            // Iterate only up to the relevant number of rows and columns
            for ($row; $row <= $lastRow; $row++) {
                $newSheetCol = 'A';
                for ($col = 'A'; $col <= 'N'; $col++) {
                    /*
                    if ($skipCol && ($col == 'C' || $col == 'H')) {
                        continue;
                    }
                    */
                    $cellValue = $sheet->getCell($col . $row)->getCalculatedValue();
                    $newSheet->setCellValue($newSheetCol . $newSheetRow, $cellValue);
                    // Copy cell style
                    $newSheet->getStyle($newSheetCol . $newSheetRow)->applyFromArray(
                        $sheet->getStyle($col . $row)->exportArray()
                    );
                    $newSheetCol++;
                }
                $newSheetRow++;
            }
            if ($sheetIndex == 0) {
                // For the first sheet, merge all rows without offset up to $lastNewSheetRow
                foreach ($sheet->getMergeCells() as $mergeCell) {
                    $mergeCellRange = Coordinate::extractAllCellReferencesInRange($mergeCell);
                    $firstCell = $mergeCellRange[0];
                    $lastCell = $mergeCellRange[count($mergeCellRange) - 1];
                    
                    $firstColumn = Coordinate::columnIndexFromString(Coordinate::coordinateFromString($firstCell)[0]);
                    $lastColumn = Coordinate::columnIndexFromString(Coordinate::coordinateFromString($lastCell)[0]);
                    
                    $firstRow = Coordinate::coordinateFromString($firstCell)[1];
                    $lastMergeRow = Coordinate::coordinateFromString($lastCell)[1];
                    
                    if ($lastMergeRow <= $lastNewSheetRow) {
                        $newMergeRange = Coordinate::stringFromColumnIndex($firstColumn) . $firstRow . ':' . 
                                         Coordinate::stringFromColumnIndex($lastColumn) . $lastMergeRow;
                        $newSheet->mergeCells($newMergeRange);
                    }
                }
            } else {
                // For subsequent sheets, map A7:M$lastRow onto A$firstNewSheetRow:M$lastNewSheetRow
                foreach ($sheet->getMergeCells() as $mergeCell) {
                    $mergeCellRange = Coordinate::extractAllCellReferencesInRange($mergeCell);
                    $firstCell = $mergeCellRange[0];
                    $lastCell = $mergeCellRange[count($mergeCellRange) - 1];
                    
                    $firstColumn = Coordinate::columnIndexFromString(Coordinate::coordinateFromString($firstCell)[0]);
                    $lastColumn = Coordinate::columnIndexFromString(Coordinate::coordinateFromString($lastCell)[0]);
                    
                    $firstRow = Coordinate::coordinateFromString($firstCell)[1];
                    $lastMergeRow = Coordinate::coordinateFromString($lastCell)[1];
                    
                    if ($firstRow >= 7 && $lastMergeRow <= $lastRow) {
                        $newFirstRow = $firstNewSheetRow + ($firstRow - 7);
                        $newLastRow = $firstNewSheetRow + ($lastMergeRow - 7);
                        
                        if ($newLastRow <= $lastNewSheetRow) {
                            $newMergeRange = Coordinate::stringFromColumnIndex($firstColumn) . $newFirstRow . ':' . 
                                             Coordinate::stringFromColumnIndex($lastColumn) . $newLastRow;
                            $newSheet->mergeCells($newMergeRange);
                        }
                    }
                }
            }
            
            if ($thisSheetWithExtraCols) {
                if ($firstNewSheetRow<9) {
                    $firstNewSheetRow = 9;
                }
                if ($sheetIndex == 1) {
                    $lastNewSheetRow -=7;
                }
                for ($tempRow = $firstNewSheetRow; $tempRow <= $lastNewSheetRow; $tempRow++) {
                    Log::info("Merging cells B" . $tempRow . ":C" . $tempRow);
                    $newSheet->mergeCells('B' . $tempRow . ':C' . $tempRow);
                    Log::info("Merging cells H" . $tempRow . ":I" . $tempRow);
                    $newSheet->mergeCells('H' . $tempRow . ':I' . $tempRow);
                }
                $lastNewSheetRow += 7;
            }
            if ($lastIteration) {
                $newSheet = $this->setRunningTotals($newSheet, $totalsBoxStart[0] . ($lastNewSheetRow-7), $sumsSection);
            }
            $sheetIndex++;
        }
    }

    

    $writer = IOFactory::createWriter($newSpreadsheet, 'Xlsx');
    $filename = $design->id . "_" . time() . "_configured";
    $newFilePath = storage_path('app/public/orders/' . $filename . '.xlsx');
    $writer->save($newFilePath);
    return $newFilePath;
}

private function checkExtraColumn($sheet) {
    $sheetsWithExtraCol = ["плита", "винты", "лента"];
    foreach ($sheetsWithExtraCol as $exception) {
        if (strpos($sheet, $exception) !== false) {
            return true;
        }
    }
    return false;
    }

    public function updateRunningTotals($sheet, $totalsBoxStart, $sumsSection) {
        for ($i=0; $i<8; $i++) {
            $columnForTotals = $totalsBoxStart[0];
            $rowForTotals = substr($totalsBoxStart, 1);
            Log::info("Updating totals for sheet " . $sheet->getTitle() . " for cell " . $columnForTotals . ($rowForTotals+$i));
            if (isset($sumsSection[$i])) {
                $sumsSection[$i] = $sumsSection[$i] + $sheet->getCell($columnForTotals . ($rowForTotals+$i))->getCalculatedValue();
            } else {
                $sumsSection[$i] = $sheet->getCell($columnForTotals . ($rowForTotals+$i))->getCalculatedValue();
            }
        }
        return $sumsSection;
    }

    public function setRunningTotals($sheet, $totalsBoxStart, $sumsSection) {
        for ($i=0; $i<8; $i++) {
            $columnForTotals = $totalsBoxStart[0];
            $rowForTotals = substr($totalsBoxStart, 1);
            $sheet->setCellValue($columnForTotals . ($rowForTotals+$i-7), $sumsSection[$i]);
        }
        return $sheet;
    }

    public function handlePriceIndexing($spreadsheet, $design)
    {
         // Reset calculation cache
        $dangerous = 0;
        Calculation::getInstance($spreadsheet)->clearCalculationCache();
        Log::info("Cleared calculation cache and now processing data for " . $design->id);
        $this->processDatasheet($spreadsheet, $design);
        $minimum = 0;
        $exceptionalSheets = ["Мягкая", "Железо", "плита", "лента", "СВ-Рост"];
        foreach ($spreadsheet->getWorksheetIterator() as $worksheet) {
            if (strpos($worksheet->getTitle(), "Смета") !== false) {
                $continue = false;
                if ($this->temp == 0 && (strpos($worksheet->getTitle(), "КС 145х45") !== false || strpos($worksheet->getTitle(), "КС 145x45") !== false)) {
                    //delete all columns beyond N
                    Log::info("Deleting columns beyond N");
                    $worksheet->removeColumn("N", 18250);
                    $this->temp = 1;
                }
                foreach ($exceptionalSheets as $exception) {
                    if (strpos($worksheet->getTitle(), $exception) !== false) {
                        $continue = true;
                        switch ($exception) {
                            case "лента":
                                $this->processLenta($spreadsheet, $design);
                                break;
                            case "СВ-Рост":
                                $this->processSVR($spreadsheet, $design);
                                break;
                            case "Железо":
                                $this->metalAndPlastic($spreadsheet, $design);
                                break;
                            case "плита":
                                $this->processPlate($spreadsheet, $design);
                                break;
                        }
                    }
                }
                if (!$continue) {
                    $variation = str_replace("Смета ", "", $worksheet->getTitle());
                    $variation_ref = Redis::get($worksheet->getTitle());
                    $labour = $worksheet->getCell("C3")->getCalculatedValue();
                    $material = $worksheet->getCell("C4")->getCalculatedValue();
                    $total = $worksheet->getCell("C5")->getCalculatedValue();
                    //Log::info("project " . $design->id . " variation " . $variation . " labour " . $labour . " material " . $material . " total " . $total);
                    $material = is_numeric($material) && !is_nan($material) ? round($material, 0) : 999;
                    $labour = is_numeric($labour) && !is_nan($labour) ? round($labour, 0) : 999;
                    $total = is_numeric($total) && !is_nan($total) ? round($total, 0) : 999;

                    $results[$design->id][$variation] = [
                        "labour" => $labour,
                        "material" => $material,
                        "total" => $total,
                    ];

                    if ($material == 999 || $labour == 999 || $total == 999) {
                        $dangerous = 1;
                    }
                    if (strpos($design->title, "ОЦБ") !== false && $variation == 'ХВ 200') {
                        Redis::set($design->id, json_encode($results[$design->id][$variation]['material']));
                    }
                    if (strpos($design->title, "ПБ") !== false && $variation == 'брус КС 145х145') {
                        Redis::set($design->id, json_encode($results[$design->id][$variation]['material']));
                    }
                    if (strpos($design->title, "ОЦБ") !== false && $variation == 'ХВ 180') {
                        Redis::set($design->id . "_seasonal", json_encode($results[$design->id][$variation]['material']));
                    }
                    if (strpos($design->title, "ПБ") !== false && $variation == 'брус КС 145х45') {
                        Redis::set($design->id . "_seasonal", json_encode($results[$design->id][$variation]['material']));
                    }

                    // Add records to Redis
                    $redisKey = $design->id . "_" . $variation_ref;
                    Redis::set($redisKey, json_encode($results[$design->id][$variation]));
                }
                $price = json_decode(Redis::get($design->id), true);
                $priceSeasonal = json_decode(Redis::get($design->id . "_seasonal"), true);
                $priceArray = ['price' => $price, 'seasonal' => $priceSeasonal];
                $details = $priceArray;
                if (!is_null($design->details)) {
                    try {
                        $details = json_decode($design->details, true);
                        if (is_array($details)) {
                        $details = array_merge($details, $priceArray);
                        } else {
                            $details = $priceArray;
                        }
                    } catch (\Exception $e) {
                        Log::error("Error merging price into details for design " . $design->id . ": " . $e->getMessage());
                        $details = $priceArray;
                    }
                }
                if ($dangerous == 1) {
                    $design->mMetrics = 1;
                    $design->save();
                } else {
                    $design->mMetrics = 0;
                }
                $design->details = json_encode($details);
                $design->save();
            }
        }
        $zeroCostSheets = ["fNone", "rNone"];
        foreach ($zeroCostSheets as $sheet) {
         $redisKey = $design->id . "_" . $sheet;
         Redis::set($redisKey, json_encode(["labour" => 0, "material" => 0, "total" => 0]));
        }
    }

    public function processLenta($spreadsheet, $design) {
        foreach ($this->fVariationArray as $variation) {
            Calculation::getInstance($spreadsheet)->clearCalculationCache();
            $worksheet = $spreadsheet->getSheetByName('data');
            $worksheet->setCellValue("D8", $variation[1]);
            $worksheet->setCellValue("D5", $variation[2]);
            $worksheet = $spreadsheet->getSheetByName('Смета лента 600х300');
            $labour = $worksheet->getCell("C3")->getCalculatedValue();
            $material = $worksheet->getCell("C4")->getCalculatedValue();
            $total = $worksheet->getCell("C5")->getCalculatedValue();
            
            $material = is_numeric($material) && !is_nan($material) ? round($material, 0) : 999;
            $labour = is_numeric($labour) && !is_nan($labour) ? round($labour, 0) : 999;
            $total = is_numeric($total) && !is_nan($total) ? round($total, 0) : 999;

            $result = [
                "labour" => $labour,
                "material" => $material,
                "total" => $total,
            ];

            // Add records to Redis
            $redisKey = $design->id . "_" . $variation[3];
            Redis::set($redisKey, json_encode($result));
        }
    }

    public function processSVR($spreadsheet, $design) {
        foreach ($this->fVariationArray as $variation) {
            Calculation::getInstance($spreadsheet)->clearCalculationCache();
            $worksheet = $spreadsheet->getSheetByName('data');
            $worksheet->setCellValue("D8", $variation[1]);
            $worksheet->setCellValue("D5", $variation[2]);
            $worksheet = $spreadsheet->getSheetByName('Смета СВ-Рост 600х300');
            $labour = $worksheet->getCell("C3")->getCalculatedValue();
            $material = $worksheet->getCell("C4")->getCalculatedValue();
            $total = $worksheet->getCell("C5")->getCalculatedValue();

            $material = is_numeric($material) && !is_nan($material) ? round($material, 0) : 999;
            $labour = is_numeric($labour) && !is_nan($labour) ? round($labour, 0) : 999;
            $total = is_numeric($total) && !is_nan($total) ? round($total, 0) : 999;

            $result = [
                "labour" => $labour,
                "material" => $material,
                "total" => $total,
            ];

            // Add records to Redis
            $redisKey = $design->id . "_" . $variation[4];
            Redis::set($redisKey, json_encode($result));
        }
    }

    public function processPlate($spreadsheet, $design) {
        $variation = $this->plitaVariationArray;
        
        foreach ($variation as $variation) {
            Calculation::getInstance($spreadsheet)->clearCalculationCache();
            $worksheet = $spreadsheet->getSheetByName('data');
            $worksheet->setCellValue("D87", $variation[0]);
        $worksheet = $spreadsheet->getSheetByName('Смета плита');
        $labour = $worksheet->getCell("C3")->getCalculatedValue();
        $material = $worksheet->getCell("C4")->getCalculatedValue();
        $total = $worksheet->getCell("C5")->getCalculatedValue();

        $material = is_numeric($material) && !is_nan($material) ? round($material, 0) : 999;
            $labour = is_numeric($labour) && !is_nan($labour) ? round($labour, 0) : 999;
            $total = is_numeric($total) && !is_nan($total) ? round($total, 0) : 999;

            $result = [
                "labour" => $labour,
                "material" => $material,
                "total" => $total,
        ];

            // Add records to Redis
            $redisKey = $design->id . "_" . $variation[1];
            Redis::set($redisKey, json_encode($result));
        }
    }

    public function metalAndPlastic($spreadsheet, $design) {
        $variation = $this->metalAndPlasticVariationArray;
        foreach ($variation as $variation) {
            $worksheet = $spreadsheet->getSheetByName($variation[0]);
            Calculation::getInstance($spreadsheet)->clearCalculationCache();
            $labour = $worksheet->getCell("C3")->getCalculatedValue();
            $labourCell = $worksheet->getCell("C3")->getValue();
            $labourCell = str_replace('=', '', $labourCell);
            $deduction1 = $worksheet->getCell($variation[2])->getCalculatedValue();
            $worksheet->setCellValue($labourCell, $labour - $deduction1);
            $material = $worksheet->getCell("C4")->getCalculatedValue();
            $materialCell = $worksheet->getCell("C4")->getValue();
            $materialCell = str_replace('=', '', $materialCell);
            $deduction2 = $worksheet->getCell($variation[3])->getCalculatedValue();
            $worksheet->setCellValue($materialCell, $material - $deduction2);
            $total = $worksheet->getCell("C5")->getCalculatedValue();
            $total = $total - $deduction1 - $deduction2;

            $material = is_numeric($material) && !is_nan($material) ? round($material, 0) : 999;
            $labour = is_numeric($labour) && !is_nan($labour) ? round($labour, 0) : 999;
            $total = is_numeric($total) && !is_nan($total) ? round($total, 0) : 999;

            $result = [
                "labour" => $labour,
                "material" => $material,
                "total" => $total,
            ];

            // Add records to Redis
            $redisKey = $design->id . "_" . $variation[1];
            Redis::set($redisKey, json_encode($result));
        }
    }
    public function processDatasheet($spreadsheet, $design)
    {
        $sheet = $spreadsheet->getSheetByName("data");
 
         //Фундамент лента
         foreach ($this->cellMappings['fLenta'] as $cell => $value) {
             $sheet->setCellValue($cell, is_string($value) ? $design->$value : $value);
         }
         //Фундамент Винта
        foreach ($this->cellMappings['fVinta'] as $cell => $value) {
            $sheet->setCellValue($cell, $design->$value);
        }

        //Брус
         
        foreach ($this->cellMappings['brus'] as $cell => $value) {
            $sheet->setCellValue($cell, $design->$value);
        }

        // floor areas for Brus
        $allFloors = $design->areafl0[0];
        foreach ($this->cellMappings['brusFloors'] as $cell => $key) {
            $sheet->setCellValue($cell, $allFloors[$key]);
        }

        // OCB and stuff
        foreach ($this->cellMappings['ocb'] as $cell => $value) {
            $sheet->setCellValue($cell, $design->$value);
        }
        Log::info("Updated OCB section of Data");

        // Кровля мягкая
        foreach ($this->cellMappings['rSoft'] as $cell => $value) {
            $sheet->setCellValue($cell, $design->$value);
        }
        Log::info("Updated rSoft section of Data");

        // PV parts
        foreach ($this->cellMappings['pv'] as $cell => $value) {
            $sheet->setCellValue($cell, $design->$value);
        }

        // MV parts
        foreach ($this->cellMappings['mv'] as $cell => $value) {
            $sheet->setCellValue($cell, $design->$value);
        }

        // srRoofSection
        foreach ($this->cellMappings['srRoofSection'] as $cell => $value) {
            $sheet->setCellValue($cell, $design->$value);
        }
        Log::info("Updated srRoofSection section of Data");

        $startingIndex = 281;
        $endingIndex = 302;
        foreach ($design->metaList as $metal) {
            $sheet->setCellValue('L' . $startingIndex, $metal['width']);
            $sheet->setCellValue('M' . $startingIndex, $metal['quantity']);
            $startingIndex++;
        }
        for ($i = $startingIndex; $i <= $endingIndex; $i++) {
            $sheet->setCellValue('L' . $i, 0);
            $sheet->setCellValue('M' . $i, 0);
        }

        Log::info("Moving to Balki");
        $sheet = $spreadsheet->getSheetByName('балки');
        $startingIndex = 15;
        $endingIndex = 40;
        // Mapping of floor names to numbers/letters
        $floorMapping = [
            "Первый" => '1', // Первый
            "Второй" => '2', // Второй
            "Третий" => '3', // Третий
            "Чердак" => 'Ч'  // Чердак
        ];
        $uniqueLengths = [];
        foreach ($design->floorsList as $room) {
            $floorNumber = $floorMapping[$room['floors']] ?? ''; 
            $sheet->setCellValue('E' . $startingIndex, $room['length']);
            if (!in_array($room['length'], $uniqueLengths)) {
                $uniqueLengths[] = $room['length'];
            }
            $sheet->setCellValue('F' . $startingIndex, $room['width']);
            $sheet->setCellValue('G' . $startingIndex, 630);
            $sheet->setCellValue('H' . $startingIndex, $floorNumber);
            $startingIndex++;
        }

        for ($i = $startingIndex; $i <= $endingIndex; $i++) {
            $sheet->setCellValue('E' . $i, "");
            $sheet->setCellValue('F' . $i, "");
            $sheet->setCellValue('G' . $i, "");
            $sheet->setCellValue('H' . $i, "");
        }
        Calculation::getInstance($spreadsheet)->clearCalculationCache();
        $startingIndex = 15;
        $endingIndex = 40;
        foreach ($uniqueLengths as $length) {
            $sheet->setCellValue('P' . $startingIndex, $length);
            $startingIndex++;
        }
        for ($i = $startingIndex; $i <= $endingIndex; $i++) {
            $sheet->setCellValue('P' . $i, "");
        }

        $sheet->setCellValue('P15', "=UNIQUE(E15:E40)");

        Log::info("Balki section completed");
        return $spreadsheet;
    }
}