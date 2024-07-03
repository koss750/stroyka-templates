<div class="col-md-4 mb-4">
    <div class="card template-box">
        <div class="card-body">
            <h3 class="card-title">{{ $title }}</h3>
            @if($template)
                <p class="template-info">{{ $template->category }} - {{ $template->name }}</p>
                <form action="{{ route('update-template', ['id' => $template->id]) }}" method="POST" enctype="multipart/form-data" class="template-form">
                    @method('PUT')
                    @csrf
                    <div class="form-group">
                        <label for="{{ $category }}_file">Выберите файл:</label>
                        <input id="{{ $category }}_file" type="file" class="form-control-file" name="file" required>
                    </div>
                    <input type="hidden" name="name" value="{{ $name }}"/>
                    <button type="submit" class="btn btn-primary">Обработать</button>
                </form>
                @if($category == 'main')
                    <script>
                        document.querySelector('.template-form').addEventListener('submit', function() {
                            document.getElementById('reindexButton').disabled = true;
                    });
                    </script>
                    <button class="btn btn-warning mt-2" id="reindexButton">Индексировать цены</button>
                    <div class="progress" id="progressBarContainer" style="display:none; margin-top: 10px; color: green;">
                        <div class="progress-bar bg-success" id="progressBar" role="progressbar" style="width: 10%; padding: 9px; color: green" aria-valuenow="0" aria-valuemin="0" aria-valuemax="100" ></div>
                    </div>
                    <script>
                        document.getElementById('reindexButton').addEventListener('click', function() {
                            document.getElementById('reindexButton').style.display = 'none'; // Ensure the progress bar container is visible when the button is clicked
                            document.getElementById('progressBarContainer').style.display = ''; // Ensure the progress bar container is visible when the button is clicked
                            document.getElementById('progressBarContainer').style.marginTop = '10px'; // Ensure the progress bar container is visible when the button is clicked
                            document.getElementById('progressBarContainer').style.padding = '5px';
                            document.getElementById('progressBarContainer').style.textAlign = 'center';
                            let progressBar = document.getElementById('progressBar');
                            let fakeProgress = [20, 50, 80, 92, 95, 98, 99, 100];
                            let delay = 1000;

                            fakeProgress.forEach((value, index) => {
                                setTimeout(() => {
                                    progressBar.style.width = value + '%';
                                    progressBar.setAttribute('aria-valuenow', value);
                                    if (value === 100) {
                                        // Replace progress bar with text when complete
                                        document.getElementById('progressBarContainer').innerHTML = 'Цены обновяться за 3 минуты';
                                    }
                                }, delay * (index + 1));
                            });

                            fetch('{{ route('reindex-prices', ['count' => 500]) }}', {
                                method: 'GET',
                                headers: {
                                    'X-Requested-With': 'XMLHttpRequest',
                                    'X-CSRF-TOKEN': '{{ csrf_token() }}'
                                }
                            }).then(response => {
                                if (response.ok) {
                                    console.log('Reindexing started successfully');
                                } else {
                                    console.error('Reindexing failed');
                                }
                            }).catch(error => {
                                console.error('Error:', error);
                            });
                        });
                    </script>
                @endif
            @else
                <p class="no-template-message">Шаблон для {{ $title }} не загружен. Оч плохо.</p>
                <form action="{{ route('store-template') }}" method="POST" enctype="multipart/form-data" class="template-form">
                    @csrf
                    <input type="hidden" name="category" value="{{ $category }}">
                    <div class="form-group">
                        <label for="{{ $category }}_file">Выберите файл:</label>
                        <input id="{{ $category }}_file" type="file" class="form-control-file" name="file" required>
                    </div>
                    <input type="hidden" name="name" value="{{ $name }}"/>
                    <button type="submit" class="btn btn-primary">Загрузить</button>
                    
                </form>
            @endif
            @if($category == 'main')
            <a class="btn btn-success mt-2" style="display: none;" href="{{ route('download-template', ['category' => $category]) }}">Скачать шаблон</a>
                <button class="btn btn-secondary mt-2" data-toggle="modal" data-target="#{{ $category }}Modal">Сгенерировать смету</button>
                
            @endif
        </div>
    </div>
</div>

<style>
    .template-box {
        height: 100%;
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        border-radius: 8px;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
        transition: transform 0.3s ease;
    }

    .template-box:hover {
        transform: translateY(-5px);
    }

    .template-box .card-title {
        font-size: 24px;
        font-weight: bold;
        margin-bottom: 10px;
    }

    .template-box .template-info {
        margin-bottom: 15px;
        color: #6c757d;
    }

    .template-box .no-template-message {
        margin-bottom: 15px;
        color: #dc3545;
    }

    .template-box .template-form {
        margin-bottom: 20px;
    }

    .template-box .form-group label {
        font-weight: bold;
    }

    .template-box .btn-primary {
        background-color: #007bff;
        border-color: #007bff;
    }

    .template-box .btn-primary:hover {
        background-color: #0069d9;
        border-color: #0062cc;
    }

    .template-box .btn-secondary {
        background-color: #6c757d;
        border-color: #6c757d;
    }

    .template-box .btn-secondary:hover {
        background-color: #5a6268;
        border-color: #545b62;
    }
</style>