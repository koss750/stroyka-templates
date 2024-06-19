<?php

use App\Http\Controllers\ProfileController;
use Illuminate\Support\Facades\Route;
use Illuminate\Http\Request;
use App\Models\Design;
use App\Http\Controllers\TemplateController;
use App\Http\Controllers\FulfillmentController;

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

require __DIR__.'/auth.php';
