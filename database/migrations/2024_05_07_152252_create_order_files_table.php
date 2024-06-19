<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    /**
     * Run the migrations.
     */
    public function up()
    {
        Schema::create('order_files', function (Blueprint $table) {
            $table->id();
            $table->string('filename');
            $table->string('sheetname');
            $table->unsignedBigInteger('design_id');
            $table->text('params')->nullable();
            $table->timestamps();

            // Foreign key to the design table
            $table->foreign('design_id')->references('id')->on('designs')->onDelete('cascade');
        });
    }

    public function down()
    {
        Schema::dropIfExists('order_files');
    }
};
