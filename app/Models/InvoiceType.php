<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class InvoiceType extends Model
{
    use HasFactory;

    protected $table = 'invoice_structures';

    protected $fillable = [
        'ref',
        'title',
        'depth',
        'parent',
        'label',
        'params',
        'sheetnames'
    ];

    protected $cast = [
        'properties' => 'array'

    ];


    // At this point, no relationships are defined as you mentioned.
}
