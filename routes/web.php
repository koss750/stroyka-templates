<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\TemplateController;


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


Route::get('/', [TemplateController::class, 'index']);
Route::post('/store-template', [TemplateController::class, 'storeTemplate'])->name('store-template');
Route::get('/store-template', [TemplateController::class, 'storeTemplate'])->name('see-template');;
Route::get('/get-template', [TemplateController::class, 'getTemplate']);