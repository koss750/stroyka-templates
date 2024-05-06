<?php

namespace App\Http\Controllers;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Illuminate\Http\Request;
use App\Models\Design;
use Illuminate\Support\Facades\Log;
use Symfony\Component\HttpFoundation\Response;


class Fulfillment extends Controller
{
    public function process(Request $request)
    {
        
        //establishing where the request is from
        if ($request->has('debug') && $request->debug > 0) {
            $debug = $request->debug;
        } else {
            
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

            $filePath = storage_path('templates/' . $filename);
            if (!file_exists($filePath)) {
                throw new \Exception("File does not exist.");
            }

            if ($debug) {
                Log::info("Loading Excel file from $filePath");
            }

            // Load the Excel file
            $spreadsheet = IOFactory::load($filePath);
            $sheet = $spreadsheet->getSheetByName('data');
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
            $sheet->setCellValue('D86', $design->vfLength);

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
            $sheet->setCellValue('D294', $design->srEndn);
            $sheet->setCellValue('D295', $design->srGvozd);
            $sheet->setCellValue('D296', $design->srSam70);
            $sheet->setCellValue('D297', $design->srPack);
            $sheet->setCellValue('D298', $design->srIzospanAM);
            $sheet->setCellValue('D299', $design->srIzospanAM35);
            $sheet->setCellValue('D300', $design->srLenta);
            $sheet->setCellValue('D301', $design->srRokvul);
            $sheet->setCellValue('D302', $design->srIzospanB);
            $sheet->setCellValue('D303', $design->srIzospanB35);
            $sheet->setCellValue('D304', $design->srPrimUgol);
            $sheet->setCellValue('D305', $design->srPrimNakl);
            $sheet->setCellValue('D306', $design->srOSB);
            $sheet->setCellValue('D308', $design->srAero);
            $sheet->setCellValue('D309', $design->srAeroSkat);
            $sheet->setCellValue('D310', $design->stropValue);

            Log::info("Updated rSoft section of Data");
            
            Log::info("Attempting to figure out Balki");
            Log::info("Getting a list of rooms and their dimensions");
            
            $sheet = $spreadsheet->getSheetByName('балки');
            $startingIndex = 15;
            foreach ($design->floorsList as $room) {
                $sheet->setCellValue('E'. $startingIndex, $room['length']);
                $sheet->setCellValue('F'. $startingIndex, $room['width']);
                $startingIndex++;
            }
            $sheet->setCellValue('P15', "=UNIQUE(E15:E40)");

            Log::info("Balki completed");


            Log::info("Attempting to locate sheet for $sheetname");
            
            if ($sheetname != "data") {
                if ($sheetname == "balki") {
                    $sheetname = "балки";
                } else {
                $invoiceTypeObject = InvoiceType::where('label', $sheetname)->firstOrFail();
                $sheetname = $invoiceTypeObject->params;
                }
            } 
            $sheet = $spreadsheet->getSheetByName($sheetname);
            if (!$sheet) {
                throw new \Exception("Sheet '{$sheetname}' not found.");
            }

            if ($debug) {
                Log::info("Preparing to copy sheet: " . $sheet->getTitle());
            }

            // Create a new Spreadsheet for the copied data
            $newSpreadsheet = new Spreadsheet();
            $newWorksheet = $newSpreadsheet->getActiveSheet();
            $newWorksheet->setTitle("smeta");

            // Copy column widths
            foreach ($sheet->getColumnIterator() as $column) {
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

            // Save the new file to a new location
            $filename = $designId . "_" . time();
            $newFilePath = storage_path('templates/orders/' . $filename . '.xlsx');
            $writer = IOFactory::createWriter($newSpreadsheet, 'Xlsx');
            $writer->save($newFilePath);

            if ($debug) {
                Log::info("New file saved to $newFilePath");
            }

            // Return the new file path
            return response()->json(['newFilePath' => $newFilePath]);

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
