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
    
    $this->setIndexingCells($template);

    return back()->with('success', 'New template uploaded successfully.');
}

    public function updateTemplate(Request $request, $id)
{
    try {
        $template = Template::findOrFail($id);
        if ($request->hasFile('file')) {
            // Delete old file if necessary and store the new file
            Storage::delete($template->file_path);
            $path = $request->file('file')->storeAs('templates', $request->input('name', 'default_filename.xlsx'));

            $template->update(['file_path' => $path, 'name' => $request->name]);
        }
        $this->setIndexingCells($template);

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

    public function setIndexingCells(Template $template)
    {
        $filename = $template->name;
        Log::info("Processing cells");
        if (!$filename) {
            return response()->json(['error' => 'Filename is required'], Response::HTTP_BAD_REQUEST);
        }

        $filePath = storage_path('app/templates/' . $filename);
        if (!file_exists($filePath)) {
            return response()->json(['error' => 'File does not exist'], Response::HTTP_BAD_REQUEST);
        }

        try {
            Log::info("Loading spreadsheet from file: $filePath");
            $spreadsheet = IOFactory::load($filePath);

            foreach ($spreadsheet->getWorksheetIterator() as $worksheet) {
                $sheetTitle = $worksheet->getTitle();
                Log::info("Processing worksheet: $sheetTitle");
                if (strpos($sheetTitle, 'Смета') !== false) {
                    $this->fillIndexingCells($worksheet);
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

    private function fillIndexingCells($worksheet)
    {
        $searchRange = 'C25:E150';
        $searchPhrase = 'Ст-ть работ';

        $foundCell = null;
        foreach ($worksheet->rangeToArray($searchRange, null, true, true) as $rowIndex => $row) {
            foreach ($row as $colIndex => $value) {
                if ($value === $searchPhrase) {
                    $foundCell = Coordinate::stringFromColumnIndex($colIndex + 7) . ($rowIndex + 25);
                    Log::info("Found phrase '$searchPhrase' at cell: $foundCell");
                    break 2;
                }
            }
        }

        if ($foundCell) {
            $worksheet->setCellValue('C3', "=$foundCell");
            Log::info("Set cell C3 to address $foundCell");

            list($col, $row) = Coordinate::coordinateFromString($foundCell);

            $cellC4 = Coordinate::stringFromColumnIndex(Coordinate::columnIndexFromString($col)) . ($row + 4);
            $worksheet->setCellValue('C4', "=$cellC4");
            Log::info("Set cell C4 to address $cellC4");

            $cellC5 = Coordinate::stringFromColumnIndex(Coordinate::columnIndexFromString($col)) . ($row + 8);
            $worksheet->setCellValue('C5', "=$cellC5");
            Log::info("Set cell C5 to address $cellC5");
             // Set font color to white
        $worksheet->getStyle('C3')->getFont()->getColor()->setARGB(Color::COLOR_WHITE);
        $worksheet->getStyle('C4')->getFont()->getColor()->setARGB(Color::COLOR_WHITE);
        $worksheet->getStyle('C5')->getFont()->getColor()->setARGB(Color::COLOR_WHITE);
        } else {
            Log::warning("Phrase '$searchPhrase' not found in the specified range.");
        }
    }
}
