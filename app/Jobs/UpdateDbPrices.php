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

class UpdateDbPrices implements ShouldQueue
{
    use Dispatchable, InteractsWithQueue, Queueable, SerializesModels;

    public function handle()
    {
        Log::info("Updating db prices");
        //get all the designs
        $designs = Design::where('active', 1)->get();
        //get price from redis for each design
        foreach ($designs as $design) {
            try {
                $price = Redis::get($design->id);
                Log::info("Price for design " . $design->id . ": " . $price);
            } catch (\Exception $e) {
                Log::error("Error getting price from redis for design " . $design->id . ": " . $e->getMessage());
                continue;
            }
            //put price into an array and then store as json
            $priceArray = ['price' => $price];
            if (!is_null($design->details)) {
                try {
                    $details = json_decode($design->details, true);
                    if (is_array($details)) {
                        $details = array_merge($details, $priceArray);
                    } else {
                        $details = $priceArray;
                    }
                } catch (\Exception $e) {
                    Log::error("Error merging price into details for design " . $design->id . ": " . $e->getMessage());
                    $details = $priceArray;
                }
            }
            $design->update(['details' => json_encode($details)]);
        }
    }
}