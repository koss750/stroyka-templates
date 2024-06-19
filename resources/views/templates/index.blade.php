<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Управление шаблонами</title>
    <link href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" rel="stylesheet">
    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.5.4/dist/umd/popper.min.js"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
</head>
<body>
    <div class="container">
        <h1 class="my-4">Управление шаблонами</h1>

        <div class="row">
            <!-- Main Template Box -->
            @include('template_box', [
                'template' => $mainTemplate,
                'category' => 'main',
                'title' => 'Главная',
                'name' => 'Главный'
            ])

            <!-- pLenta Template Box -->
            @include('template_box', [
                'template' => $pLenta,
                'category' => 'pLenta',
                'title' => 'Ленточный с плитой',
                'name' => 'Ленточный с плитой'
            ])

            <!-- fLenta Template Box -->
            @include('template_box', [
                'template' => $fLenta,
                'category' => 'fLenta',
                'title' => 'Ленточный фундамент',
                'name' => 'Ленточный фундамент'
            ])

            <!-- plita Template Box -->
            @include('template_box', [
                'template' => $plita,
                'category' => 'plita',
                'title' => 'Расчет плиты',
                'name' => 'Расчет плиты'
            ])

            <!-- srs Template Box -->
            @include('template_box', [
                'template' => $srs,
                'category' => 'srs',
                'title' => 'Свайно-растверковый с плитой перекрытия',
                'name' => 'Свайно-растверковый с плитой перекрытия'
            ])

            <!-- sr Template Box -->
            @include('template_box', [
                'template' => $sr,
                'category' => 'sr',
                'title' => 'Свайно-растверковый',
                'name' => 'Свайно-растверковый'
            ])
        </div>

<!-- Main Template Modal -->
<!-- Main Template Modal -->
<div class="modal fade" id="mainModal" tabindex="-1" role="dialog" aria-labelledby="mainModalLabel" aria-hidden="true">
    <div class="modal-dialog" role="document">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="mainModalLabel">Сгенерировать смету (Главная)</h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                </button>
            </div>
            <div class="modal-body">
                <form action="/external" method="GET">
                <div class="form-group">
                    <label for="designInput">Название проекта:</label>
                    <input type="text" class="form-control" id="designInput" required>
                    <input type="hidden" name="design" id="designId">
                </div>
                    <div class="form-group">
                        <label for="variant">Размер ленты:</label>
                        <input type="text" class="form-control" name="variant" value="600x300" required>
                    </div>
                    <div class="form-group">
                        <input type="hidden" class="form-control" name="sheetname" value="all">
                    </div>
                    <div class="form-group form-check">
                        <input type="checkbox" class="form-check-input" name="labour">
                        <label class="form-check-label" for="labour">Включить цены за работы</label>
                    </div>
                    <input type="hidden" name="filename" value="{{ $mainTemplate->name }}">
                    <input type="hidden" name="debug" value="1"/>
                    <button type="submit" class="btn btn-primary">Сгенерировать</button>
                </form>
            </div>
        </div>
    </div>
</div>

<!-- pLenta Template Modal -->
<div class="modal fade" id="pLentaModal" tabindex="-1" role="dialog" aria-labelledby="pLentaModalLabel" aria-hidden="true">
    <div class="modal-dialog" role="document">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="pLentaModalLabel">Сгенерировать смету (Ленточный с плитой)</h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                </button>
            </div>
            <div class="modal-body">
                <form action="/external" method="GET">
                    <div class="form-group">
                        <label for="design">Номер проекта:</label>
                        <input type="text" class="form-control" name="design" required>
                    </div>
                    <!-- Add pLenta template-specific form fields here -->
                    <input type="hidden" name="category" value="pLenta">
                    <input type="hidden" name="debug" value="1"/>
                    <button type="submit" class="btn btn-primary">Сгенерировать</button>
                </form>
            </div>
        </div>
    </div>
</div>

<!-- fLenta Template Modal -->
<div class="modal fade" id="fLentaModal" tabindex="-1" role="dialog" aria-labelledby="fLentaModalLabel" aria-hidden="true">
    <div class="modal-dialog" role="document">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="fLentaModalLabel">Сгенерировать смету (Ленточный фундамент)</h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                </button>
            </div>
            <div class="modal-body">
                <form action="/external" method="GET">
                    <div class="form-group">
                        <label for="design">Номер проекта:</label>
                        <input type="text" class="form-control" name="design" required>
                    </div>
                    <!-- Add fLenta template-specific form fields here -->
                    <input type="hidden" name="category" value="fLenta">
                    <input type="hidden" name="debug" value="1"/>
                    <button type="submit" class="btn btn-primary">Сгенерировать</button>
                </form>
            </div>
        </div>
    </div>
</div>

<!-- plita Template Modal -->
<div class="modal fade" id="plitaModal" tabindex="-1" role="dialog" aria-labelledby="plitaModalLabel" aria-hidden="true">
    <div class="modal-dialog" role="document">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="plitaModalLabel">Сгенерировать смету (Расчет плиты)</h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                </button>
            </div>
            <div class="modal-body">
                <form action="/external" method="GET">
                    <div class="form-group">
                        <label for="design">Номер проекта:</label>
                        <input type="text" class="form-control" name="design" required>
                    </div>
                    <!-- Add plita template-specific form fields here -->
                    <input type="hidden" name="category" value="plita">
                    <input type="hidden" name="debug" value="1"/>
                    <button type="submit" class="btn btn-primary">Сгенерировать</button>
                </form>
            </div>
        </div>
    </div>
</div>

<!-- srs Template Modal -->
<div class="modal fade" id="srsModal" tabindex="-1" role="dialog" aria-labelledby="srsModalLabel" aria-hidden="true">
    <div class="modal-dialog" role="document">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="srsModalLabel">Сгенерировать смету (Свайно-растверковый с плитой перекрытия)</h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                </button>
            </div>
            <div class="modal-body">
                <form action="/external" method="GET">
                    <div class="form-group">
                        <label for="design">Номер проекта:</label>
                        <input type="text" class="form-control" name="design" required>
                    </div>
                    <!-- Add srs template-specific form fields here -->
                    <input type="hidden" name="category" value="srs">
                    <input type="hidden" name="debug" value="1"/>
                    <button type="submit" class="btn btn-primary">Сгенерировать</button>
                </form>
            </div>
        </div>
    </div>
</div>
<!-- sr Template Modal -->
<div class="modal fade" id="srModal" tabindex="-1" role="dialog" aria-labelledby="srModalLabel" aria-hidden="true">
    <div class="modal-dialog" role="document">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="srModalLabel">Сгенерировать смету (Свайно-растверковый)</h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                </button>
            </div>
            <div class="modal-body">
                <form action="/external" method="GET">
                    <div class="form-group">
                        <label for="design">Номер проекта:</label>
                        <input type="text" class="form-control" name="design" required>
                    </div>
                    <!-- Add sr template-specific form fields here -->
                    <input type="hidden" name="category" value="sr">
                    <input type="hidden" name="debug" value="1"/>
                    <button type="submit" class="btn btn-primary">Сгенерировать</button>
                </form>
            </div>
        </div>
    </div>
</div>

    </div>
</body>
</html>

<script>
    document.addEventListener('DOMContentLoaded', function() {
        const designInput = document.getElementById('designInput');
        const designIdInput = document.getElementById('designId');
        const messageElement = document.createElement('div');
        messageElement.classList.add('message');
        designInput.parentNode.insertBefore(messageElement, designInput.nextSibling);

        designInput.addEventListener('input', function() {
            const designTitle = this.value;
            if (designTitle) {
                // Make a request to fetch the project ID based on the title
                fetch('/get-project-id?title=' + encodeURIComponent(designTitle))
                    .then(response => response.json())
                    .then(data => {
                        if (data.success) {
                            designIdInput.value = data.id;
                            messageElement.textContent = 'Проект найден';
                            messageElement.style.color = 'green';
                        } else {
                            designIdInput.value = '';
                            messageElement.textContent = 'Проект не найден';
                            messageElement.style.color = 'red';
                        }
                    })
                    .catch(error => {
                        console.error('Error:', error);
                        messageElement.textContent = 'Ошибка при поиске проекта';
                        messageElement.style.color = 'red';
                    });
            } else {
                designIdInput.value = '';
                messageElement.textContent = '';
            }
        });
    });
</script>
