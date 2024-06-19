<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class OrderFile extends Model
{
    use HasFactory;
    protected $fillable = ['filename', 'sheetname', 'design_id', 'params'];
    public function design()
    {
        return $this->belongsTo(Design::class);
    }
}
