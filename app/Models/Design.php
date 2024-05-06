<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Template extends Model
{
    use HasFactory;
    protected $table = 'designs';
	
	public $timestamps = true;

    // Fillable fields for mass assignment
    protected $fillable = [
        'details',
        'created_at',
        'updated_at',
        'category',
        'size',
        'mMetrics',
        'length',
        'width',
        'code',
        'numOrders',
        'materialType',
        'floorsList',
        'baseType',
        'roofType',
        'roofSquare',
        'mainSquare',
        'baseLength',
        'baseD20',
        'baseD20F',
        'baseD20Rub',
        'baseD20RubF',
        'baseBalk1',
        'baseBalkF',
        'baseBalk2',
        'wallsOut',
        'wallsIn',
        'wallsPerOut',
        'wallsPerIn',
        'rubRoof',
        'skatList',
        'krovlaList',
        'stropList',
        'stropValue',
        'endovList',
        'areafl0',
        'metaList',
        // ... include other fields as needed
    ];

	protected $casts = [
	    'floorsList' => 'json',
	    'category' => 'json',
	    'skatList' => 'json',
	    'stropList' => 'json',
	    'areafl0' => 'json',
        'endovList' => 'json',
        'metaList' => 'json',
        'krovlaList' => 'json',
        'pvParts' => 'json', // or 'object' if you prefer
        'mvParts' => 'json' // or 'object' if you prefer
    ];
}