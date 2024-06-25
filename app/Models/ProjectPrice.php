<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use App\Http\Controllers\FulfilmentController;

class ProjectPrice extends Model
{
    use HasFactory;

    protected $fillable = [
        'design_id',
        'invoice_type_id',
        'price',
    ];

    /**
     * Get the design associated with the project price.
     */
    public function design()
    {
        return $this->belongsTo(Design::class);
    }

    /**
     * Get the invoice type associated with the project price.
     */
    public function invoiceType()
    {
        return $this->belongsTo(InvoiceType::class);
    }
}
