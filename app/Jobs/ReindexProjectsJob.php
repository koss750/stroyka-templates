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
use App\Services\SpreadsheetService;

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
        $service = new SpreadsheetService();
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
            
            $service->handle($filePath, $latestProjects, true);

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

    
}
