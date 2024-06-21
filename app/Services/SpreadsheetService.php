<?php

namespace App\Services;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Illuminate\Support\Facades\Log;
use Illuminate\Support\Facades\Redis;
use PhpOffice\PhpSpreadsheet\Calculation\Calculation;
use App\Models\Design;
use Illuminate\Support\Collection;

class SpreadsheetService
{
    public $temp = 0;

    private $cellMappings;

    public function __construct()
    {
        $this->cellMappings = [
            'fLenta' => [
                'D4' => 'lfLength',
                'D5' => 0.6,
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
                'D289' => 'srKonShir',
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
    }

    public function handle($filePath, $design=1, $multiple=false, $debug=1) {
        
        try {
            $spreadsheet = IOFactory::createReader('Xlsx')->load($filePath);
        } catch (\Exception $e) {
            throw $e;
        }
        if ($multiple && $design instanceof Collection) {
            $designs = $design; //actually is a collection
            foreach ($designs as $design) {
                $this->handlePriceIndexing($spreadsheet, $design);
            }
        } else if ($multiple && !$design instanceof Collection) {
            throw new \Exception("Design is not a collection");
        } else {
            $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
            $spreadsheet = $this->handlePriceIndexing($spreadsheet, $design);
            $filename = $design->id . "_" . time();
            $newFilePath = storage_path('app/public/orders/' . $filename . '.xlsx');
            $writer->save($newFilePath);
            return $newFilePath;
        }
    }

    public function handlePriceIndexing($spreadsheet, $design)
    {
         // Reset calculation cache
        Calculation::getInstance($spreadsheet)->clearCalculationCache();
        $this->processDatasheet($spreadsheet, $design);
        $minimum = 0;
        foreach ($spreadsheet->getWorksheetIterator() as $worksheet) {
            if (strpos($worksheet->getTitle(), "Смета") !== false) {
                if ($this->temp == 0 && (strpos($worksheet->getTitle(), "КС 145х45") !== false || strpos($worksheet->getTitle(), "КС 145x45") !== false)) {
                    //delete all columns beyond N
                    Log::info("Deleting columns beyond N");
                    $worksheet->removeColumn("N", 18250);
                    $this->temp = 1;
                }
                $variation = str_replace("Смета ", "", $worksheet->getTitle());
                $variation_ref = Redis::get($worksheet->getTitle());
                $labour = $worksheet->getCell("C3")->getCalculatedValue();
                $material = $worksheet->getCell("C4")->getCalculatedValue();
                $total = $worksheet->getCell("C5")->getCalculatedValue();

                $material = is_numeric($material) && !is_nan($material) ? round($material, 0) : 999;
                $labour = is_numeric($labour) && !is_nan($labour) ? round($labour, 0) : 999;
                $total = is_numeric($total) && !is_nan($total) ? round($total, 0) : 999;

                $results[$design->id][$variation] = [
                    "labour" => $labour,
                    "material" => $material,
                    "total" => $total,
                ];

                if ($variation == 'Мягкая' || $variation == 'ХВР 200') {
                    $minimum = $material + $minimum;
                }

                // Add records to Redis
                $redisKey = $design->id . "_" . $variation_ref;
                Redis::set($redisKey, json_encode($results[$design->id][$variation]));
            }
        }
        Redis::set("{$design->id}", round($minimum, 0));
    }

    public function processDatasheet($spreadsheet, $design)
    {
        $sheet = $spreadsheet->getSheetByName("data");
 
         //Фундамент лента
         foreach ($this->cellMappings['fLenta'] as $cell => $value) {
             $sheet->setCellValue($cell, is_string($value) ? $design->$value : $value);
         }
         Log::info("Updated fLenta section of Data");
 
         //Фундамент Винта
         foreach ($this->cellMappings['fVinta'] as $cell => $value) {
             $sheet->setCellValue($cell, $design->$value);
         }
        Log::info("Updated fVinta section of Data");

        //Брус
         
        foreach ($this->cellMappings['brus'] as $cell => $value) {
            $sheet->setCellValue($cell, $design->$value);
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
            $startingIndex++;
        }

        for ($i = $startingIndex; $i <= $endingIndex; $i++) {
            $sheet->setCellValue('E' . $i, "");
            $sheet->setCellValue('F' . $i, "");
            $sheet->setCellValue('G' . $i, "");
            $sheet->setCellValue('H' . $i, "");
        }

        $sheet->setCellValue('P15', "=UNIQUE(E15:E40)");

        Log::info("Balki completed");
        return $spreadsheet;
    }
}