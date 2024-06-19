<?php

namespace App\Jobs;

use App\Models\Design;
use App\Models\OrderFile;
use Illuminate\Bus\Queueable;
use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Foundation\Bus\Dispatchable;
use Illuminate\Queue\InteractsWithQueue;
use Illuminate\Queue\SerializesModels;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use Illuminate\Support\Facades\Log;
use Symfony\Component\HttpFoundation\Response;
use Illuminate\Support\Facades\Redis;
use App\Models\InvoiceType;
use App\Jobs\UpdateDbPrices;
use PhpOffice\PhpSpreadsheet\Calculation\Calculation;
use Illuminate\Support\Facades\Cache;

class ReindexProjectsJob implements ShouldQueue
{
    use Dispatchable, InteractsWithQueue, Queueable, SerializesModels;

    protected $projectCount;
    protected $temp = 0;

    /**
     * Create a new job instance.
     *
     * @return void
     */
    public function __construct($projectCount)
    {
        $this->projectCount = $projectCount;
    }

    /**
     * Execute the job.
     *
     * @return void
     */
    public function handle()
    {
        Log::info("Starting ReindexProjectsJob", [
            "project_count" => $this->projectCount,
        ]);

        try {
            $latestProjects = Design::orderBy("created_at", "desc")
                ->where('active', 1)
                ->take($this->projectCount)
                ->get();
            Log::info("Fetched latest projects", [
                "count" => $latestProjects->count(),
            ]);

            $results = [];

            $filename = "Главный";
            $filePath = storage_path("app/templates/" . $filename);
            if (!file_exists($filePath)) {
                throw new \Exception("File does not exist.");
            }

            // Get all InvoiceType objects and store their sheetnames and labels in Redis
            $invoiceTypeObject = InvoiceType::all();
            foreach ($invoiceTypeObject as $invoiceType) {
                Redis::set($invoiceType->sheetname, $invoiceType->label);
            }
            Log::info("Loading Excel file from $filePath");
            try {
                $spreadsheet = IOFactory::createReader('Xlsx')->load($filePath);
            } catch (\Exception $e) {
                throw $e;
            }
            foreach ($latestProjects as $design) {
                Log::info("Processing project " . $design->id);
                $this->handleProject($spreadsheet, $design, $filePath, $results);
                //$this->debug($spreadsheet, $design, $filePath, $results);
            }

            Log::info("Finished processing all projects");
            dispatch(new UpdateDbPrices());
            Log::info("Dispatched UpdateDbPricesJob");

        } catch (\Exception $e) {
            Log::error("Error in ReindexProjectsJob", [
                "message" => $e->getMessage(),
            ]);
            throw $e;
        }
    }

    private function debug($spreadsheet, $design, $filePath, &$results) {
        $sheet = $spreadsheet->getSheetByName("data");
        Log::info("design->lfLength: " . $design->lfLength);
        $sheet->setCellValue("A1", $design->lfLength);
        
        //output A1,A2,A3 values
        Log::info("A1: " . $sheet->getCell("A1")->getValue());
        Log::info("A2: " . $sheet->getCell("A2")->getCalculatedValue());
        Log::info("A3: " . $sheet->getCell("A3")->getCalculatedValue());

    }

    private function handleProject($spreadsheet, $design, $filePath, &$results)
    {
        // Reset calculation cache
        Calculation::getInstance($spreadsheet)->clearCalculationCache();
        $tapeWidth = 0.6;
        $tapeLength = 0.5;
        $sheet = $spreadsheet->getSheetByName("data");
        Log::info("Updating cells for $design->lfLength");
        $sheet->setCellValue("D4", $design->lfLength);
        $sheet->setCellValue("D5", $tapeWidth);
        $sheet->setCellValue("D8", $tapeLength);
        $sheet->setCellValue("D9", $design->lfAngleX);
        $sheet->setCellValue("D10", $design->lfAngleT);
        $sheet->setCellValue("D11", $design->lfAngleG);
        $sheet->setCellValue("D12", $design->lfAngle45);
        $sheet->setCellValue("D14", 0.2);
        $sheet->setCellValue("D15", 0.2);
        $sheet->setCellValue("D16", $design->mfSquare);

        // Фундамент Винта
        $sheet->setCellValue("D44", $design->vfCount);
        $sheet->setCellValue("D86", $design->vfLength);

        // Брус
        $sheet->setCellValue("C113", $design->baseD20RubF);
        $sheet->setCellValue("C114", $design->baseBalk1);
        $sheet->setCellValue("C115", $design->stolby);

        // Floor areas for Brus
        $allFloors = $design->areafl0[0];
        $sheet->setCellValue("I112", $allFloors["Sfl0"]);
        $sheet->setCellValue("I113", $allFloors["Sfl1"]);
        $sheet->setCellValue("I114", $allFloors["Sfl2"]);
        $sheet->setCellValue("I115", $allFloors["Sfl3"]);
        $sheet->setCellValue("I116", $allFloors["Sfl4"]);
        $sheet->setCellValue("D283", $design->roofSquare);
        $sheet->setCellValue("D284", $design->srCherep);
        $sheet->setCellValue("D285", $design->srKover);
        $sheet->setCellValue("D286", $design->srKonK);
        $sheet->setCellValue("D287", $design->srMastika1);
        $sheet->setCellValue("D288", $design->srMastika);
        $sheet->setCellValue("D289", $design->srKonShir);
        $sheet->setCellValue("D290", $design->srKonOneSkat);
        $sheet->setCellValue("D291", $design->srPlanVetr);
        $sheet->setCellValue("D292", $design->srPlanK);
        $sheet->setCellValue("D293", $design->srKapelnik);
        $sheet->setCellValue("D294", $design->srEndn);
        $sheet->setCellValue("D295", $design->srGvozd);
        $sheet->setCellValue("D296", $design->srSam70);
        $sheet->setCellValue("D297", $design->srPack);
        $sheet->setCellValue("D298", $design->srIzospanAM);
        $sheet->setCellValue("D299", $design->srIzospanAM35);
        $sheet->setCellValue("D300", $design->srLenta);
        $sheet->setCellValue("D301", $design->srRokvul);
        $sheet->setCellValue("D302", $design->srIzospanB);
        $sheet->setCellValue("D303", $design->srIzospanB35);
        $sheet->setCellValue("D304", $design->srPrimUgol);
        $sheet->setCellValue("D305", $design->srPrimNakl);
        $sheet->setCellValue("D306", $design->srOSB);
        $sheet->setCellValue("D308", $design->srAero);
        $sheet->setCellValue("D309", $design->srAeroSkat);
        $sheet->setCellValue("D310", $design->stropValue);

        $sheet = $spreadsheet->getSheetByName("балки");
        $startingIndex = 15;
        foreach ($design->floorsList as $room) {
            $sheet->setCellValue("E" . $startingIndex, $room["length"]);
            $sheet->setCellValue("F" . $startingIndex, $room["width"]);
            $startingIndex++;
        }
        $sheet->setCellValue("P15", "=UNIQUE(E15:E40)");
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
}
