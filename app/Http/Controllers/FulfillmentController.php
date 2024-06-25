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
use App\Services\SpreadsheetService;

class FulfillmentController extends Controller
{
    protected $spreadsheetService;

    public function __construct(SpreadsheetService $spreadsheetService)
    {
        $this->spreadsheetService = $spreadsheetService;
    }

    public function processLatestProjects($projectCount)
    {
        //execution time limit to 10min
        ini_set('max_execution_time', 600);
        ReindexProjectsJob::dispatch($projectCount);
        return response()->json(['message' => 'Reindexing job dispatched'], 200);
    }

    public function createSmeta($design, $config)
    {
        $filePath = storage_path("app/templates/Главный");
        $design = Design::find($design);
        return $this->spreadsheetService->handle($filePath, $design, true, false, 1, $config);
    }

    public function process(Request $request)
    {
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

        $testVals = [
            // Define your test values here
        ];

        $filename = $request->filename;
        $filePath = storage_path("app/templates/" . $filename);
        if (!file_exists($filePath)) {
            throw new \Exception("File does not exist.");
        }
        $sheetname = $request->sheetname;
        $cellData = $request->cellData ?? $testVals;
        if ($sheetname == 'all') {
            $filePath = $this->spreadsheetService->handle($filePath, $design, false);
            return response()->download($filePath);
        } else {
            $filePath = $this->spreadsheetService->handle($filePath, $design, false, true);
            return response()->download($filePath);
        }
    }
}