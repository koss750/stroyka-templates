<?php

namespace App\Http\Controllers;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Illuminate\Http\Request;
use App\Models\Design;
use Illuminate\Support\Facades\Log;
use Symfony\Component\HttpFoundation\Response;
use App\Jobs\ReindexProjectsJob;
use App\Models\InvoiceType;


class FulfillmentController extends Controller
{
    public function processLatestProjects($projectCount)
    {
        //execution time limit to 10min
        ini_set('max_execution_time', 600);
        ReindexProjectsJob::dispatch($projectCount);
        return response()->json(['message' => 'Reindexing job dispatched'], 200);
    }

    public function process(Request $request)
    {
        
        //establishing where the request is from
        if ($request->has('debug') && $request->debug > 0) {
            $debug = $request->debug;
        } else {
            $debug = 1;
        }
        // Check if debugging is enabled
        

        $designId = $request->design;
        $variant = $request->variant;
        $design = Design::where('id', $designId)->firstOrFail();
        $tapeWidth = $variant[4];
        if ($tapeWidth == 1) {
            $tapeWidth = 10;
        }
        $tapeWidth = $tapeWidth / 10;
        $tapeLength = $tapeWidth - 0.1;

        try {
            $testVals = [
                // Define your test values here
            ];

            $filename = $request->filename;
            $sheetname = $request->sheetname;
            $cellData = $request->cellData ?? $testVals; // Expecting an array of ['cell' => 'value']

            if ($debug) {
                Log::info("Received filename: $filename, sheetname: $sheetname");
                Log::info("Cell data: " . json_encode($cellData));
            }

            if (!$filename || !$sheetname || !is_array($cellData)) {
                throw new \Exception("Invalid input parameters.");
            }

            $filePath = storage_path('app/templates/' . $filename);
            if (!file_exists($filePath)) {
                Log::info("File does not exist: $filePath");
                throw new \Exception("File does not exist.");
            }

            if ($debug) {
                Log::info("Loading Excel file from $filePath");
            }

            // Load the Excel file
            $spreadsheet = IOFactory::load($filePath);
            $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
            $sheet = $spreadsheet->getSheetByName('data');
            Log::info("Value of D283 in 'data' sheet: " . $spreadsheet->getSheetByName('data')->getCell('D283')->getCalculatedValue());
            if (!$sheet) {
                throw new \Exception("Sheet '{$sheetname}' not found.");
            }

            if ($debug) {
                Log::info("Loaded sheet: " . $sheet->getTitle());
            }

            // Update the cells in the specified sheet
            foreach ($cellData as $cell => $value) {
                $sheet->setCellValue($cell, $value);
                if ($debug) {
                    Log::info("Updated cell $cell with value $value");
                }
            }
            
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

            //if sheetname is "all", save and download the new spreadsheet
            if ($sheetname == "all") {
                $filename = $designId . "_" . time();
                $newFilePath = storage_path('app/public/orders/' . $filename . '.xlsx');
                $writer->save($newFilePath);
                // Return the new file
                return response()->download($newFilePath);
            }
            

            // Create a new Spreadsheet for the copied data
            $newSpreadsheet = new Spreadsheet();
            $newWorksheet = $newSpreadsheet->getActiveSheet();

            //loop through all sheets in the spreadsheet
            foreach ($spreadsheet->getWorksheetIterator() as $sheet) {
                // Copy column widths for columns A:N
                foreach ($sheet->getColumnIterator('A', 'N') as $column) {
                    $columnIndex = $column->getColumnIndex();
                    $newWorksheet->getColumnDimension($columnIndex)
                    ->setWidth($sheet->getColumnDimension($columnIndex)->getWidth());
                if ($debug) {
                    Log::info("Copied column width for column $columnIndex");
                    }
                }

                // Copy merged cells
                foreach ($sheet->getMergeCells() as $mergeRange) {
                    $newWorksheet->mergeCells($mergeRange);
                    if ($debug) {
                        Log::info("Copied merged cell range $mergeRange");
                    }
                }

            // Copy values and formatting (but no formulas) from the original sheet
            foreach ($sheet->getRowIterator() as $row) {
                $rowIndex = $row->getRowIndex();
                foreach ($sheet->getColumnIterator() as $column) {
                    $columnIndex = $column->getColumnIndex();
                    try {
                        $cell = $sheet->getCell($columnIndex . $rowIndex);
                        $style = $sheet->getStyle($columnIndex . $rowIndex)->exportArray();
                        $newWorksheet->setCellValue($cell->getCoordinate(), $cell->getCalculatedValue());
                        $newWorksheet->getStyle($columnIndex . $rowIndex)->applyFromArray($style);
                        if ($debug==2) {
                            //Log::info("Copied cell {$columnIndex}{$rowIndex}");
                        }
                        } catch (\Exception $e) {
                            Log::error("Error copying cell {$columnIndex}{$rowIndex}: " . $e->getMessage());
                        }
                    }
                }
            }

            

            // Save the new file to a new location
            $filename = $designId . "_" . time();
            $newFilePath = storage_path('templates/orders/' . $filename . '.xlsx');
            $publicpath = public_path('orders/' . $filename . '.xlsx');
            IOFactory::save($spreadsheet, $newFilePath);

            return response()->download($publicpath);

        } catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
            Log::error("Spreadsheet read error: " . $e->getMessage());
            return response()->json(['error' => 'Error reading the spreadsheet'], Response::HTTP_BAD_REQUEST);
        } catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
            Log::error("Spreadsheet error: " . $e->getMessage());
            return response()->json(['error' => 'Spreadsheet processing error'], Response::HTTP_BAD_REQUEST);
        } catch (\Exception $e) {
            Log::error("General error: " . $e->getMessage());
            return response()->json(['error' => $e->getMessage()], Response::HTTP_BAD_REQUEST);
        }
    }
}