<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Carbon\Carbon;
use App\Http\Controllers\FulfillmentController as FC;
use App\Services\SpreadsheetService;

class Project extends Model
{
    use HasFactory;

    protected $fillable = [
        'user_id',
        'ip_address',
        'payment_reference',
        'payment_amount',
        'design_id',
        'selected_configuration',
        'filepath',
    ];

    protected $casts = [
        'selected_configuration' => 'array',
        'payment_amount' => 'decimal:2',
    ];

    public function user()
    {
        return $this->belongsTo(User::class);
    }

    public function design()
    {
        return $this->belongsTo(Design::class);
    }

    public static function boot()
    {
        parent::boot();
        
        static::creating(function ($project) {
            if (!$project->payment_reference) {
                $project->payment_reference = Carbon::now()->format('YmdHis');
            }
        });
    }

    public function getFormattedCreatedAtAttribute()
    {
        return $this->created_at->format('Y-m-d H:i:s');
    }

    public function getFormattedUpdatedAtAttribute()
    {
        return $this->updated_at->format('Y-m-d H:i:s');
    }

    public function createSmeta()
    {
        $SS = new SpreadsheetService();
        $FC = new FC($SS);
        $link = $FC->createSmeta($this->design_id, $this->selected_configuration);
        $this->update(['filepath' => $link]);
    }
}