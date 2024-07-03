<?php

namespace App\Http\Controllers;

use App\Models\Template;
use Illuminate\Http\Request;
use App\Models\InvoiceType;
use App\Models\OrderFile;
use Illuminate\Support\Facades\Storage;
use Illuminate\Support\Facades\Log;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Symfony\Component\HttpFoundation\Response;
use PhpOffice\PhpSpreadsheet\Spreadsheet;    
use PhpOffice\PhpSpreadsheet\Style\Color;

class TemplateController extends Controller
{
    protected $validPasscode = '123';

    /**
     * Store the template file.
     *
     * @param \Illuminate\Http\Request $request
     * @return \Illuminate\Http\JsonResponse
     */
    public function storeTemplate(Request $request)
{
    $validatedData = $request->validate([
        'file' => 'required|file',
        'category' => 'required|string'
    ]);

    $path = $request->file('file')->store('templates');
    $name = $request->file('file')->getClientOriginalName();

    $template = Template::create([
        'name' => $name,
        'file_path' => $path,
        'category' => $validatedData['category']
    ]);
    
    $this->initialProcessing($template);

    return back()->with('success', 'New template uploaded successfully.');
}

    public function updateTemplate(Request $request, $id)
{
    try {
        $template = Template::findOrFail($id);
        if ($request->hasFile('file')) {
            // Delete old file if necessary and store the new file
            Storage::delete($template->file_path);
            $path = $request->file('file')->storeAs('templates', $request->input('name', 'default_filename') . '.xlsx'); 

            $template->update(['file_path' => $path, 'name' => $request->name]);
        }
        $this->initialProcessing($template);

        return back()->with('success', 'Template updated successfully.');
    } catch (\Exception $e) {
        return back()->with('error', 'Error updating template: ' . $e->getMessage());
    }
}

    /**
     * Retrieve the template file.
     *
     * @param \Illuminate\Http\Request $request
     * @return \Illuminate\Http\Response
     */
    public function getTemplate(Request $request)
    {
        try {
            // Validate the incoming request
            $validatedData = $request->validate([
                'template_name' => 'required|string',
                'passcode' => 'required|string',
            ]);

            // Verify passcode
            if ($validatedData['passcode'] !== $this->validPasscode) {
                return response()->json(['error' => 'Unauthorized'], 401);
            }

            // Retrieve the template from the database
            $template = Template::where('name', $validatedData['template_name'])->first();

            if (!$template) {
                return response()->json(['error' => 'Template not found'], 404);
            }

            $path = $template->file_path;

            if (!Storage::exists($path)) {
                return response()->json(['error' => 'File not found on disk'], 404);
            }

            return response()->download(storage_path('app/' . $path));
        } catch (\Exception $e) {
            return response()->json(['error' => 'Failed to retrieve template', 'message' => $e->getMessage()], 500);
        }
    }

    public function index()
    {
        // Retrieve each template safely, or return null if not found
        $mainTemplate = Template::where('category', 'main')->first();
        $sr = Template::where('category', 'sr')->first();
        $srs = Template::where('category', 'srs')->first();
        $plita = Template::where('category', 'plita')->first();
        $fLenta = Template::where('category', 'flenta')->first();  // Match the case
        $pLenta = Template::where('category', 'plenta')->first();  // Match the case
        $templates = Template::all();
        $orderFiles = OrderFile::with('design')->latest()->take(5)->get();
    
        return view('templates.index', compact('mainTemplate', 'pLenta', 'fLenta', 'plita', 'sr', 'srs', 'orderFiles', 'templates'));
    }

    public function initialProcessing(Template $template)
    {
        $filename = $template->name . '.xlsx';
        Log::info("Processing cells");
        if (!$filename) {
            return response()->json(['error' => 'Filename is required'], Response::HTTP_BAD_REQUEST);
        }

        $filePath = storage_path('app/templates/' . $filename);
        if (!file_exists($filePath)) {
            return response()->json(['error' => 'File does not exist'], Response::HTTP_BAD_REQUEST);
        }
        Log::info("File exists");
        try {
            Log::info("Loading spreadsheet from file: $filePath");
            $spreadsheet = IOFactory::load($filePath);

            foreach ($spreadsheet->getWorksheetIterator() as $worksheet) {
                $sheetTitle = $worksheet->getTitle();
                Log::info("Processing worksheet: $sheetTitle");
                if (strpos($sheetTitle, 'Смета') !== false) {
                    $this->initialSheetProcessing($worksheet);
                }
            }

            $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
            $writer->save($filePath);

            Log::info("File processed and saved to: $filePath");
            return response()->json(['message' => 'File processed successfully', 'path' => "/storage/templates/$filename"], Response::HTTP_OK);

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

    private function initialSheetProcessing($worksheet)
    {
        $dbSheetSpec = [
            "total" => [],
            "index_total_start" => "total_1",
            "index_delivery_start" => "DS",
            "index_delivery_end" => "DE",
            "index_smeta_alt_start" => "AE",
            "index_smeta_alt_end" => "DE+1d",
            "labour_cost_col" => "J",
            "material_cost_col" => "K",
            "sections" => [
                0 => [
                    "start" => "index_smeta_alt_start+1/2",
                    "end" => "null",
                ]
            ]
        ];
        $searchRange = 'A6:M150';
        $startingRow = 6;
        $searchPhrase = 'Ст-ть работ';
        $foundCell = null;
        $cleanUpRows = false;
        $sectionCounter = 0;
        $cleanUpColumn = 'J';
        $keywords = [
            'Итого работы' => 'SE',
            '  Прием материалов. Транспортные расходы	' => 'D',
            'Аренда а/м до 1,5 тн' => 'DL',
        ];
        foreach ($worksheet->rangeToArray($searchRange, null, true, true) as $rowIndex => $row) {
            foreach ($row as $colIndex => $value) {
                $coorCollIndex = $colIndex+1;
                if ($rowIndex < 5 && $colIndex < 2) {
                    if (strpos($value, 'смета') != false || strpos($value, 'СМЕТА') != false) {
                        Log::info("found смета at " . Coordinate::stringFromColumnIndex($coorCollIndex) . ($rowIndex+$startingRow) . " with value: $value");
                        $dbSheetSpec['index_smeta_alt_start'] = $rowIndex+$startingRow;
                        continue;
                    }
                    continue;
                } elseif ($rowIndex < 3) {
                    continue;
                }
                if ($value == 11 && $rowIndex < 9) {
                    $lastCell = Coordinate::stringFromColumnIndex($colIndex) . ($rowIndex+$startingRow);
                    $lastCell = $lastCell[0];
                    switch ($lastCell) {
                        case 'K':
                            $dbSheetSpec['labour_cost_col'] = "E";
                            $dbSheetSpec['material_cost_col'] = "J";
                            break;
                        case 'L':
                            $dbSheetSpec['labour_cost_col'] = "F";
                            $dbSheetSpec['material_cost_col'] = "K";
                            break;
                        case 'M':
                            $dbSheetSpec['labour_cost_col'] = "G";
                            $dbSheetSpec['material_cost_col'] = "L";
                            break;
                    }
                    //looking for first section
                    //value of next row
                    $nextRow = $worksheet->getCell('A' . ($rowIndex+$startingRow+2))->getValue();
                    if (strpos($nextRow, '1. ') !== false) {
                        $dbSheetSpec['sections'][0]['title'] = $nextRow;
                        $dbSheetSpec['sections'][0]['start'] = $rowIndex+$startingRow+2;
                    }
                    $nextRow = $worksheet->getCell('A' . ($rowIndex+$startingRow+1))->getValue();
                    if (strpos($nextRow, '1. ') !== false) {
                        $dbSheetSpec['sections'][0]['title'] = $nextRow;
                        $dbSheetSpec['sections'][0]['start'] = $rowIndex+$startingRow+1;
                    }
                }
                $cell = Coordinate::stringFromColumnIndex($colIndex) . ($rowIndex);
                if ($value === $searchPhrase) {
                    $foundCell = Coordinate::stringFromColumnIndex($coorCollIndex+6) . ($rowIndex+$startingRow);
                    Log::info("Found phrase '$searchPhrase' at cell: $foundCell");
                    break 2;
                }
                if (isset($keywords[$value])) {
                    Log::info("Found phrase '$value' at cell: $cell");
                    $worksheet->setCellValue('N' . ($rowIndex+$startingRow), $keywords[$value]);
                    if ($value === 'Итого работы' && $cleanUpRows) {
                        $dbSheetSpec['index_smeta_alt_end'] = $rowIndex+$startingRow+1;
                        $dbSheetSpec['index_delivery_end'] = $rowIndex+$startingRow;
                    }
                    switch ($value) {
                        case 'Итого работы':
                            $dbSheetSpec['sections'][$sectionCounter]['end'] = $rowIndex+$startingRow;
                            $sectionCounter++;
                            $dbSheetSpec['sections'][$sectionCounter]['start'] = $rowIndex+$startingRow+1;
                            $dbSheetSpec['sections'][$sectionCounter]['title'] = $worksheet->getCell('A' . ($rowIndex+$startingRow+1))->getValue();

                            //$worksheet->setCellValue('N' . ($rowIndex+$startingRow+1), 'SS');
                            //check if next 2 rows A col has "материалов" in it with strpos
                            if (strpos($worksheet->getCell('A' . ($rowIndex+$startingRow+1))->getValue(), 'материалов') !== false) {
                                $dbSheetSpec['index_delivery_start'] = $rowIndex+$startingRow+1;
                                //$worksheet->setCellValue('O' . ($rowIndex+$startingRow+1), 'DS');
                                $cleanUpRows = $rowIndex+$startingRow+1;
                                break 2;
                            }
                            if (strpos($worksheet->getCell('A' . ($rowIndex+$startingRow+2))->getValue(), 'материалов') !== false) {
                                $dbSheetSpec['index_delivery_start'] = $rowIndex+$startingRow+2;
                                //$worksheet->setCellValue('O' . ($rowIndex+$startingRow+2), 'DS');
                                $cleanUpRows = $rowIndex+$startingRow+2;
                                break 2;
                            }
                            $cleanUpRows = 0;
                            break;
                    }
                    
                }
            }
        }
        unset($dbSheetSpec['sections'][$sectionCounter]);
        //Removing penultimate line in the summary box
        //Saving formulas of summary box along with other important sheet info to the DB
        if ($foundCell) {
            list($col, $row) = Coordinate::coordinateFromString($foundCell);
            $worksheet->removeRow($row + $startingRow, 2);
            $cellC3 = Coordinate::stringFromColumnIndex(Coordinate::columnIndexFromString($col)) . ($row);
            $worksheet->setCellValue('C3', "=$cellC3");
            Log::info("Set cell C3 to address $cellC3");

            $cellC4 = Coordinate::stringFromColumnIndex(Coordinate::columnIndexFromString($col)) . ($row + 4);
            $worksheet->setCellValue('C4', "=$cellC4");
            Log::info("Set cell C4 to address $cellC4");

            $deliveryCost = Coordinate::stringFromColumnIndex(Coordinate::columnIndexFromString($col)) . ($row + 6);
            Log::info("Set cell $deliveryCost to address $deliveryCost");

            $cellC5 = Coordinate::stringFromColumnIndex(Coordinate::columnIndexFromString($col)) . ($row + 8);
            $worksheet->setCellValue('C5', "=$cellC5" . "-" . $deliveryCost);
            Log::info("Set cell C5 to address $cellC5");
             // Set font color to white
            $worksheet->getStyle('C3')->getFont()->getColor()->setARGB(Color::COLOR_WHITE);
            $worksheet->getStyle('C4')->getFont()->getColor()->setARGB(Color::COLOR_WHITE);
            $worksheet->getStyle('C5')->getFont()->getColor()->setARGB(Color::COLOR_WHITE);
            $col = Coordinate::stringFromColumnIndex(Coordinate::columnIndexFromString($col)-2);
            $total1cell = $col . ($row);
            $total2cell = $col . ($row+1);
            $total3cell = $col . ($row+2);
            $total4cell = $col . ($row+3);
            $total5cell = $col . ($row+4);
            $total6cell = $col . ($row+5);
            $total7cell = $col . ($row+6);
            $total8cell = $col . ($row+7);
            $dbSheetSpec['total'][0] = $worksheet->getCell($total1cell)->getValue();
            $total1Calculated = $worksheet->getCell("C3")->getValue();
            $dbSheetSpec['index_total_start'] = substr($total1Calculated, 1);
            $dbSheetSpec['total'][1] = $worksheet->getCell($total2cell)->getValue();
            $dbSheetSpec['total'][2] = $worksheet->getCell($total3cell)->getValue();
            $dbSheetSpec['total'][3] = $worksheet->getCell($total4cell)->getValue();
            $dbSheetSpec['total'][4] = $worksheet->getCell($total5cell)->getValue();
            $dbSheetSpec['total'][5] = $worksheet->getCell($total6cell)->getValue();
            $dbSheetSpec['total'][6] = $worksheet->getCell($total7cell)->getValue();
            $invoiceType = InvoiceType::where('sheetname', $worksheet->getTitle())->get();
            foreach ($invoiceType as $type) {
                $type->sheet_spec = $dbSheetSpec;
                $type->save();
            }
        } else {
            Log::warning("Phrase '$searchPhrase' not found in the specified range.");
        }
    }

    public function downloadTemplate($category)
    {
        $template = Template::where('category', $category)->first();
        return Storage::download($template->file_path);
    }
}
