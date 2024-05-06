<?php

namespace App\Http\Controllers;

use App\Models\Template;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Storage;

class TemplateController extends Controller
{
    protected $validPasscode = '123';

    /**
     * Store the template file.
     *
     * @param \Illuminate\Http\Request $request
     * @return \Illuminate\Http\JsonResponse
     */
    public function storeTemplate(Request $request)
    {
        try {
            // Validate the incoming request
            $validatedData = $request->validate([
                'template_name' => 'required|string',
                'file' => 'required|file',
                'passcode' => 'required|string',
            ]);

            // Verify passcode
            if ($validatedData['passcode'] !== $this->validPasscode) {
                return response()->json(['error' => 'Unauthorized'], 401);
            }

            // Store the file
            $path = $validatedData['file']->storeAs('templates', $validatedData['template_name']);

            // Store record in the database
            $template = Template::create([
                'name' => $validatedData['template_name'],
                'file_path' => $path,
            ]);

            return response()->json(['path' => $template->file_path], 201);
        } catch (\Exception $e) {
            return response()->json(['error' => 'Failed to store template', 'message' => $e->getMessage()], 500);
        }
    }

    /**
     * Retrieve the template file.
     *
     * @param \Illuminate\Http\Request $request
     * @return \Illuminate\Http\Response
     */
    public function getTemplate(Request $request)
    {
        try {
            // Validate the incoming request
            $validatedData = $request->validate([
                'template_name' => 'required|string',
                'passcode' => 'required|string',
            ]);

            // Verify passcode
            if ($validatedData['passcode'] !== $this->validPasscode) {
                return response()->json(['error' => 'Unauthorized'], 401);
            }

            // Retrieve the template from the database
            $template = Template::where('name', $validatedData['template_name'])->first();

            if (!$template) {
                return response()->json(['error' => 'Template not found'], 404);
            }

            $path = $template->file_path;

            if (!Storage::exists($path)) {
                return response()->json(['error' => 'File not found on disk'], 404);
            }

            return response()->download(storage_path('app/' . $path));
        } catch (\Exception $e) {
            return response()->json(['error' => 'Failed to retrieve template', 'message' => $e->getMessage()], 500);
        }
    }

    public function index()
    {
        $templates = Template::all();
        return view('templates.index', ['templates' => $templates]);
    }
}
