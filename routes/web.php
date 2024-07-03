<?php

use App\Http\Controllers\ProfileController;
use Illuminate\Support\Facades\Route;
use Illuminate\Http\Request;
use App\Models\Design;
use App\Http\Controllers\TemplateController;
use App\Http\Controllers\FulfillmentController;
use Illuminate\Support\Facades\DB;
use App\Models\Project;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider and all of them will
| be assigned to the "web" middleware group. Make something great!
|
*/

Route::get('/', [TemplateController::class, 'index'])->middleware('auth');
Route::get('/external', [FulfillmentController::class, 'process']);
Route::get('/reindex-prices/{count}', [FulfillmentController::class, 'processLatestProjects'])->name('reindex-prices');
Route::post('/store-template', [TemplateController::class, 'storeTemplate'])->name('store-template');
Route::put('/update-template/{id}', [TemplateController::class, 'updateTemplate'])->name('update-template');
Route::get('/get-template', [TemplateController::class, 'getTemplate']);
Route::get('/download-template/{category}', [TemplateController::class, 'downloadTemplate'])->name('download-template');

Route::middleware('auth')->group(function () {
    Route::get('/profile', [ProfileController::class, 'edit'])->name('profile.edit');
    Route::patch('/profile', [ProfileController::class, 'update'])->name('profile.update');
    Route::delete('/profile', [ProfileController::class, 'destroy'])->name('profile.destroy');
});

Route::get('/get-project-title', function (Request $request) {
    $id = $request->query('id');
    // Query the Designs table to get the project title based on the $id
    $project = Design::find($id);
    if ($project) {
        return response()->json(['success' => true, 'title' => $project->title]);
    } else {
        return response()->json(['success' => false]);
    }
});

Route::get('/get-project-id', function (Request $request) {
    $title = $request->query('title');
    // Query the database to find the project ID based on the title
    $project = DB::table('designs')->where('title', $title)->first();
    if ($project) {
        return response()->json(['success' => true, 'id' => $project->id]);
    } else {
        return response()->json(['success' => false]);
    }
});

// New route to check for sheetname in the invoice_structures table
Route::get('/get-sheetname', function (Request $request) {
    $name = $request->query('name');
    // Query the database to find the sheetname
    $sheet = DB::table('invoice_structures')->where('sheetname', $name)->first();
    if ($sheet) {
        return response()->json(['success' => true, 'name' => $sheet->sheetname]);
    } else {
        return response()->json(['success' => false]);
    }
});

// New route to get sheetname suggestions
Route::get('/get-sheetname-suggestions', function (Request $request) {
    $query = $request->query('query');
    // Query the database to find matching sheetnames
    $suggestions = DB::table('invoice_structures')
        ->where('sheetname', 'like', '%' . $query . '%')
        ->orWhere('label', 'like', '%' . $query . '%')
        ->pluck('sheetname');
    return response()->json(['success' => true, 'suggestions' => $suggestions]);
});

Route::get('/process-order/{id}', function ($id) {
    
    $order = Project::find($id);
    $order->createSmeta($order->selected_configuration);
    return response()->download($order->filepath);
});

require __DIR__.'/auth.php';

