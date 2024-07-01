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
                'D16' => 'mfSquare'
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
        foreach ($config as $sheetName => $sheetConfig) {
            $sheet = InvoiceType::where('ref', $sheetConfig)->firstOrFail();
            $sheetsToCombine[] = [$sheet->label, $sheet->sheetname];
        }
        return $sheetsToCombine;
    }

    private function processConfiguredSheets($spreadsheet, $design, $config)
{
    Calculation::getInstance($spreadsheet)->clearCalculationCache();
    $this->processDatasheet($spreadsheet, $design);

    $newSpreadsheet = new Spreadsheet();
    $newSpreadsheet->removeSheetByIndex(0); // Remove default sheet

    $sheetsToCombine = $this->getSheetsToCombine($config);

    $sheetIndex = 0;
    $newSheet = $newSpreadsheet->createSheet($sheetIndex);
    $newSheet->setTitle("Смета");
    $newSheetRow = 1;
    
    foreach ($sheetsToCombine as $sheetName) {
        $sheetTitle = $sheetName[1];
        $sheet = $spreadsheet->getSheetByName($sheetTitle);
        if ($spreadsheet->sheetNameExists($sheetTitle)) {
            $lastRow = $sheet->getCell('C3')->getValue();
            $lastRow = substr($lastRow, 2)-1;
            $skipCol = $this->checkExtraColumn($sheetTitle);
            if ($sheetIndex == 0) {
                $row = 1;
            } else $row = 8;
            // Iterate only up to the relevant number of rows and columns
            for ($row; $row <= $lastRow; $row++) {
                $newSheetCol = 'A';
                for ($col = 'A'; $col <= 'N'; $col++) {
                    if ($skipCol && ($col == 'C' || $col == 'H')) {
                        continue;
                    }
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
                // Apply merged cells only for the first $lastRow rows
                foreach ($sheet->getMergeCells() as $mergeCell) {
                    $mergeCellRange = Coordinate::extractAllCellReferencesInRange($mergeCell);
                    $firstCell = $mergeCellRange[0];
                    $lastCell = $mergeCellRange[count($mergeCellRange) - 1];
                    
                    $firstColumn = Coordinate::columnIndexFromString(Coordinate::coordinateFromString($firstCell)[0]);
                    $lastColumn = Coordinate::columnIndexFromString(Coordinate::coordinateFromString($lastCell)[0]);
                    
                    $firstRow = Coordinate::coordinateFromString($firstCell)[1];
                    $lastMergeRow = Coordinate::coordinateFromString($lastCell)[1];
                    
                    if ($lastMergeRow <= $lastRow) {
                        $offset = $newSheetRow - $row;
                        $newFirstCell = Coordinate::stringFromColumnIndex($firstColumn) . ($firstRow + $offset);
                        $newLastCell = Coordinate::stringFromColumnIndex($lastColumn) . ($lastMergeRow + $offset);
                        $newMergeRange = $newFirstCell . ':' . $newLastCell;
                        $newSheet->mergeCells($newMergeRange);
                    }
                }
            
                // Copy column dimensions
                foreach ($sheet->getColumnDimensions() as $colDim) {
                    if ($skipCol && ($colDim->getColumnIndex() == 'C' || $colDim->getColumnIndex() == 'H')) {
                        continue;
                    }
                    $newSheet->getColumnDimension($colDim->getColumnIndex())
                        ->setWidth($colDim->getWidth());
                }
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
    $sheetsWithExtraCol = ["СВ-Рост", "плита", "лента"];
    foreach ($sheetsWithExtraCol as $exception) {
        if (strpos($sheet, $exception) !== false) {
            Log::info("Extra column found for " . $sheet);
            return true;
        }
    }
    Log::info("No extra column found for " . $sheet);
    return false;
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

                    if ($variation == 'брус КС 145х145') {
                        Redis::set($design->id, json_encode($results[$design->id][$variation]['material']));
                    }
                    if ($variation == 'брус КС 145х45') {
                        Redis::set($design->id . "_seasonal", json_encode($results[$design->id][$variation]['material']));
                    }

                    // Add records to Redis
                    $redisKey = $design->id . "_" . $variation_ref;
                    Redis::set($redisKey, json_encode($results[$design->id][$variation]));
                }
                if ($dangerous == 1) {
                    $design->mMetrics = 1;
                    $design->save();
                } else {
                    $design->mMetrics = 0;
                    $design->save();
                }
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
            $worksheet->setCellValue("D5", $variation[1]);
            $worksheet->setCellValue("D8", $variation[2]);
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
            Calculation::getInstance($spreadsheet)->clearCalculationCache();
            $worksheet = $spreadsheet->getSheetByName($variation[0]);
            $labour = $worksheet->getCell("C3")->getCalculatedValue();
            $deduction1 = $worksheet->getCell($variation[2])->getCalculatedValue();
            $labour = $labour - $deduction1;
            $material = $worksheet->getCell("C4")->getCalculatedValue();
            $deduction2 = $worksheet->getCell($variation[3])->getCalculatedValue();
            $material = $material - $deduction2;
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
             Log::info("Updated cell " . $cell . " with value " . $design->$value);
         }
         Log::info("Updated fLenta section of Data");
 
         //Фундамент Винта
        foreach ($this->cellMappings['fVinta'] as $cell => $value) {
            $sheet->setCellValue($cell, $design->$value);
            Log::info("Updated cell " . $cell . " with value " . $design->$value);
        }
        Log::info("Updated fVinta section of Data");

        //Брус
         
        foreach ($this->cellMappings['brus'] as $cell => $value) {
            $sheet->setCellValue($cell, $design->$value);
            Log::info("Updated cell " . $cell . " with value " . $design->$value);
        }
        Log::info("Updated Brus section of Data");

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

        foreach ($design->floorsList as $room) {
            $floorNumber = $floorMapping[$room['floors']] ?? ''; 
            $sheet->setCellValue('E' . $startingIndex, $room['length']);
            $sheet->setCellValue('F' . $startingIndex, $room['width']);
            $sheet->setCellValue('G' . $startingIndex, 630);
            $sheet->setCellValue('H' . $startingIndex, $floorNumber);
            Log::info("Updated cell " . 'E' . $startingIndex . " with value " . $room['length']);
            Log::info("Updated cell " . 'F' . $startingIndex . " with value " . $room['width']);
            Log::info("Updated cell " . 'G' . $startingIndex . " with value " . 630);
            Log::info("Updated cell " . 'H' . $startingIndex . " with value " . $floorNumber);
            $startingIndex++;
        }

        for ($i = $startingIndex; $i <= $endingIndex; $i++) {
            $sheet->setCellValue('E' . $i, "");
            $sheet->setCellValue('F' . $i, "");
            $sheet->setCellValue('G' . $i, "");
            $sheet->setCellValue('H' . $i, "");
            Log::info("Updated cell " . 'E' . $i . " with value " . "");
            Log::info("Updated cell " . 'F' . $i . " with value " . "");
            Log::info("Updated cell " . 'G' . $i . " with value " . "");
            Log::info("Updated cell " . 'H' . $i . " with value " . "");
        }

        $sheet->setCellValue('P15', "=UNIQUE(E15:E40)");

        Log::info("Balki completed");
        return $spreadsheet;
    }
}