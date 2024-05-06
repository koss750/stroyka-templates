<!-- resources/views/templates/index.blade.php -->

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Template Management</title>
</head>
<body>
    <h1>Template Management</h1>

    <!-- List Existing Templates -->
    <h2>Existing Templates</h2>
    <table border="1">
        <thead>
            <tr>
                <th>ID</th>
                <th>Name</th>
                <th>Added On</th>
            </tr>
        </thead>
        <tbody>
            @foreach ($templates as $template)
                <tr>
                    <td>{{ $template->id }}</td>
                    <td>{{ $template->name }}</td>
                    <td>{{ $template->created_at }}</td>
                </tr>
            @endforeach
        </tbody>
    </table>

    <!-- Form to Upload New Template -->
    <h2>Upload New Template</h2>
    <form action="{{ route('store-template') }}" method="POST" enctype="multipart/form-data">
        @csrf
        <label for="template_name">Template Name:</label>
        <input type="text" name="template_name" required><br><br>

        <label for="file">File:</label>
        <input type="file" name="file" required><br><br>

        <label for="passcode">Passcode:</label>
        <input type="text" name="passcode" value="123" required><br><br>

        <button type="submit">Upload Template</button>
    </form>
</body>
</html>
