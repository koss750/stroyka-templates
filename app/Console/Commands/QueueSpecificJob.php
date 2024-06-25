<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;

class QueueSpecificJob extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'app:job {job}';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Add a specific job to the queue';

    public function handle()
    {
        $jobName = $this->argument('job');
        $jobClass = "App\\Jobs\\$jobName";

        if (!class_exists($jobClass)) {
            $this->error("Job class $jobClass does not exist.");
            return 1;
        }

        dispatch(new $jobClass());

        $this->info("Job $jobName has been added to the queue.");
    }
}
