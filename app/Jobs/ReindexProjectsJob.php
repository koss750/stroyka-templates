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
        //Фундамент лента

        $sheet->setCellValue('D4', $design->lfLength);
        $sheet->setCellValue('D5', $tapeWidth);
        $sheet->setCellValue('D8', $tapeLength);
        $sheet->setCellValue('D9', $design->lfAngleX);
        $sheet->setCellValue('D10', $design->lfAngleT);
        $sheet->setCellValue('D11', $design->lfAngleG);
        $sheet->setCellValue('D12', $design->lfAngle45);
        $sheet->setCellValue('D14', 0.2);
        $sheet->setCellValue('D15', 0.2);
        $sheet->setCellValue('D16', $design->mfSquare);

        Log::info("Updated fLenta section of Data");

        //Фундамент Винта

        $sheet->setCellValue('D44', $design->vfCount);
        $sheet->setCellValue('D45', $design->vfBalk);
        $sheet->setCellValue('D86', $design->outer_p); // Updated
        $sheet->setCellValue('D89', $design->lfAngleG); // Updated

        //Брус

        $sheet->setCellValue('C113', $design->baseD20RubF);
        $sheet->setCellValue('C114', $design->baseBalk1);
        $sheet->setCellValue('C115', $design->stolby);

        //floor areas for Brus

        $allFloors = $design->areafl0[0];
        $sheet->setCellValue('I112', $allFloors["Sfl0"]);
        $sheet->setCellValue('I113', $allFloors["Sfl1"]);
        $sheet->setCellValue('I114', $allFloors["Sfl2"]);
        $sheet->setCellValue('I115', $allFloors["Sfl3"]);
        $sheet->setCellValue('I116', $allFloors["Sfl4"]);

        Log::info("Updated Brus section of Data");

        //OCB and stuff
        $sheet->setCellValue('D169', $design->baseLength); // New
        $sheet->setCellValue('D170', $design->baseD20);

        Log::info("Updated OCB section of Data");

        // Кровля мягкая

        $sheet->setCellValue('D283', $design->roofSquare);
        $sheet->setCellValue('D284', $design->srCherep);
        $sheet->setCellValue('D285', $design->srKover);
        $sheet->setCellValue('D286', $design->srKonK);
        $sheet->setCellValue('D287', $design->srMastika1);
        $sheet->setCellValue('D288', $design->srMastika);
        $sheet->setCellValue('D289', $design->srKonShir);
        $sheet->setCellValue('D290', $design->srKonOneSkat);
        $sheet->setCellValue('D291', $design->srPlanVetr);
        $sheet->setCellValue('D292', $design->srPlanK);
        $sheet->setCellValue('D293', $design->srKapelnik);
        $sheet->setCellValue('D294', $design->mrEndv);
        $sheet->setCellValue('D295', $design->srGvozd);
        $sheet->setCellValue('D296', $design->mrSam70);
        $sheet->setCellValue('D297', $design->mrPack);
        $sheet->setCellValue('D298', $design->mrIzospanAM);
        $sheet->setCellValue('D299', $design->mrIzospanAM35);
        $sheet->setCellValue('D300', $design->mrLenta);
        $sheet->setCellValue('D301', $design->mrRokvul);
        $sheet->setCellValue('D302', $design->mrIzospanB);
        $sheet->setCellValue('D303', $design->mrIzospanB35);
        $sheet->setCellValue('D304', $design->mrPrimUgol);
        $sheet->setCellValue('D305', $design->mrPrimNakl);
        $sheet->setCellValue('D306', $design->srOSB);
        $sheet->setCellValue('D308', $design->srAero);
        $sheet->setCellValue('D309', $design->srAeroSkat);
        $sheet->setCellValue('D310', $design->stropValue);

        Log::info("Updated rSoft section of Data");

        $sheet->setCellValue('D351', $design->pvPart1);
        $sheet->setCellValue('D352', $design->pvPart2);
        $sheet->setCellValue('D353', $design->pvPart3);
        $sheet->setCellValue('D354', $design->pvPart4);
        $sheet->setCellValue('D355', $design->pvPart5);
        $sheet->setCellValue('D356', $design->pvPart6);
        $sheet->setCellValue('D357', $design->pvPart7);
        $sheet->setCellValue('D358', $design->pvPart8);
        $sheet->setCellValue('D359', $design->pvPart9);
        $sheet->setCellValue('D360', $design->pvPart10);
        $sheet->setCellValue('D361', $design->pvPart11);
        $sheet->setCellValue('D362', $design->pvPart12);
        $sheet->setCellValue('D363', $design->pvPart13);

        $sheet->setCellValue('D367', $design->mvPart1);
        $sheet->setCellValue('D368', $design->mvPart2);
        $sheet->setCellValue('D369', $design->mvPart3);
        $sheet->setCellValue('D370', $design->mvPart4);
        $sheet->setCellValue('D371', $design->mvPart5);
        $sheet->setCellValue('D372', $design->mvPart6);
        $sheet->setCellValue('D373', $design->mvPart7);
        $sheet->setCellValue('D374', $design->mvPart8);
        $sheet->setCellValue('D375', $design->mvPart9);
        $sheet->setCellValue('D376', $design->mvPart10);
        $sheet->setCellValue('D377', $design->mvPart11);
        $sheet->setCellValue('D378', $design->mvPart12);
        $sheet->setCellValue('D379', $design->mvPart13);

        Log::info("Updated PV and MV section of Data");

        // Set metaList cells
        $startingRow = 281;
        $endingRow = 302;
        foreach ($design->metaList as $item) {
            $sheet->setCellValue('L' . $startingRow, $item['width']);
            $sheet->setCellValue('M' . $startingRow, $item['quantity']);
            $startingRow++;
        } 
        for ($i = $startingRow; $i <= $endingRow; $i++) {
            $sheet->setCellValue('L' . $i, "");
            $sheet->setCellValue('M' . $i, "");
        }

        Log::info("Updated metaList section of Data");

        // Set srRoofSection cells
        $sheet->setCellValue('L306', $design->srKonShir);
        $sheet->setCellValue('L307', $design->srKonOneSkat);
        $sheet->setCellValue('L311', $design->srEndn);
        $sheet->setCellValue('L312', $design->srEndv);
        $sheet->setCellValue('L313', $design->mrSam35);
        $sheet->setCellValue('L314', $design->srSam70);
        $sheet->setCellValue('L315', $design->srPack);
        $sheet->setCellValue('L316', $design->srIzospanAM);
        $sheet->setCellValue('L317', $design->srIzospanAM35);
        $sheet->setCellValue('L322', $design->srPrimUgol);
        $sheet->setCellValue('L323', $design->srPrimNakl);

        Log::info("Updated srRoofSection section of Data. Data sheet completed. Moving to Balki");
        
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
