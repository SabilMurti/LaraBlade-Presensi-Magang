<x-app-layout>
    <x-slot name="header">
        <div class="flex items-center justify-between">
            <h2 class="text-2xl font-bold leading-tight text-base-content">
                {{ __('Data Karyawan Magang') }}
            </h2>

            <div class="flex gap-2">
                {{-- Tombol Export --}}
                <a href="{{ route('admin.karyawan.export') }}"
                    class="btn btn-success btn-md rounded-full shadow-md hover:shadow-lg transition-all duration-200 no-loading">
                    <i class="ri-download-cloud-line mr-1"></i>
                    Export Data
                </a>
                {{-- Tombol Import --}}
                <label class="btn btn-info btn-md rounded-full shadow-md hover:shadow-lg transition-all duration-200"
                    for="import_modal">
                    <i class="ri-upload-cloud-line mr-1"></i>
                    Import Data
                </label>
                <label class="btn btn-primary btn-md rounded-full shadow-md hover:shadow-lg transition-all duration-200"
                    for="create_modal">
                    <i class="ri-add-line mr-1"></i>
                    Tambah Data
                </label>
            </div>
        </div>
    </x-slot>
    {{-- Simpan data departemen di elemen tersembunyi --}}
<script id="departemen-data" type="application/json">
    @json($departemen)
</script>

<script>
// Ambil data departemen dari elemen HTML
var departemenData = [];
try {
    var elem = document.getElementById('departemen-data');
    if (elem) {
        departemenData = JSON.parse(elem.textContent);
    }
} catch(e) {
    console.error('Error parsing departemen data:', e);
}

// Fungsi delete
window.delete_button = function(nik, nama) {
    window.deleteNik = nik;
    var deleteMsg = '<p class="text-base-content text-opacity-80 mb-2">Anda yakin ingin menghapus data karyawan magang ini?</p>';
    deleteMsg += '<div class="divider text-base-content text-opacity-60"></div>';
    deleteMsg += '<div class="text-center">';
    deleteMsg += '<p class="font-bold text-lg">' + nama + '</p>';
    deleteMsg += '<p class="text-sm text-base-content text-opacity-70">NIK: ' + nik + '</p>';
    deleteMsg += '</div>';
    document.getElementById('delete_message').innerHTML = deleteMsg;
    document.getElementById('modal_delete').checked = true;
}

// Fungsi edit
window.edit_button = function(nik) {
    var loading = '<span class="loading loading-dots loading-md text-primary"></span>';
    $("#loading_edit1").html(loading);
    $("#loading_edit2").html(loading);
    $("#loading_edit3").html(loading);
    $("#loading_edit4").html(loading);
    $("#loading_edit5").html(loading);
    $("#loading_edit6").html(loading);
    $("#loading_edit7").html(loading);

    $("select[id='departemen_id']").children().remove().end();
    $(".foto-edit-preview").hide().attr("src", "");

    $.ajax({
        type: "get",
        url: "{{ route('admin.karyawan.edit') }}",
        data: {
            "_token": "{{ csrf_token() }}",
            "nik": nik
        },
        success: function(data) {
            var items = [];
            $.each(data, function(key, val) {
                items.push(val);
            });

            $("input[name='nik_lama']").val(items[0]);
            $("input[name='nik']").val(items[0]);
            $("input[name='nama_lengkap']").val(items[2]);
            $("input[name='jabatan']").val(items[4]);
            $("input[name='telepon']").val(items[5]);
            $("input[name='email']").val(items[6]);

            var options = '<option disabled>Pilih Instansi!</option>';
            for (var i = 0; i < departemenData.length; i++) {
                var item = departemenData[i];
                var isSelected = item.id == items[1] ? 'selected' : '';
                options += '<option value="' + item.id + '" ' + isSelected + '>' + item.nama + '</option>';
            }
            $("select[id='departemen_id']").html(options);

            if (items[3] != null) {
                var fotoUrl = "{{ asset('storage/unggah/karyawan') }}/" + items[3];
                $(".foto-edit-preview").attr("src", fotoUrl).show();
            } else {
                $(".foto-edit-preview").hide();
            }

            loading = "";
            $("#loading_edit1").html(loading);
            $("#loading_edit2").html(loading);
            $("#loading_edit3").html(loading);
            $("#loading_edit4").html(loading);
            $("#loading_edit5").html(loading);
            $("#loading_edit6").html(loading);
            $("#loading_edit7").html(loading);
        }
    });
}

// Preview gambar untuk create
function previewImage() {
    var image = document.querySelector('#foto');
    var imgPreview = document.querySelector('.img-preview');
    if (image.files && image.files[0]) {
        imgPreview.style.display = 'block';
        var oFReader = new FileReader();
        oFReader.readAsDataURL(image.files[0]);
        oFReader.onload = function(oFREvent) {
            imgPreview.src = oFREvent.target.result;
        }
    }
}

// Preview gambar untuk edit
function previewImageEdit() {
    var image = document.querySelector('#foto_edit');
    var imgPreview = document.querySelector('.foto-edit-preview');
    if (image.files && image.files[0]) {
        imgPreview.style.display = 'block';
        var oFReader = new FileReader();
        oFReader.readAsDataURL(image.files[0]);
        oFReader.onload = function(oFREvent) {
            imgPreview.src = oFREvent.target.result;
        }
    } else {
        imgPreview.style.display = 'none';
        imgPreview.src = '';
    }
}

// Modal alert
function showModalAlert(title, message, type) {
    type = type || 'info';
    var modalTitle = document.getElementById('modal_title');
    var modalMessage = document.getElementById('modal_message');
    var modalBox = document.querySelector('#modal_alert + .modal > .modal-box');
    
    // Cek apakah elemen ada
    if (!modalTitle || !modalMessage || !modalBox) {
        console.error('Modal elements not found');
        return;
    }
    
    modalTitle.textContent = title;
    modalMessage.innerHTML = message;
    
    if (type === 'success') {
        modalTitle.className = 'font-bold text-lg text-success';
        modalBox.className = 'modal-box border border-success bg-base-100 text-base-content';
    } else if (type === 'error') {
        modalTitle.className = 'font-bold text-lg text-error';
        modalBox.className = 'modal-box border border-error bg-base-100 text-base-content';
    } else {
        modalTitle.className = 'font-bold text-lg text-primary';
        modalBox.className = 'modal-box border border-primary bg-base-100 text-base-content';
    }
    
    document.getElementById('modal_alert').checked = true;
}

// Tunggu sampai DOM siap
document.addEventListener('DOMContentLoaded', function() {
    
    // Notifikasi session - DIPINDAHKAN KE SINI
    @if (session()->has('success'))
        showModalAlert('Berhasil!', '{{ session("success") }}', 'success');
    @endif

    @if (session()->has('error'))
        showModalAlert('Gagal!', '{!! str_replace(["'", "\r", "\n"], ["\'", "", "<br>"], session("error")) !!}', 'error');
    @endif

    // Handler hapus data
    var confirmBtn = document.getElementById('confirm_delete_btn');
    if (confirmBtn) {
        confirmBtn.addEventListener('click', function() {
            if (!window.deleteNik) return;
            
            $.ajax({
                type: "post",
                url: "{{ route('admin.karyawan.delete') }}",
                data: {
                    "_token": "{{ csrf_token() }}",
                    "nik": window.deleteNik
                },
                success: function(response) {
                    document.getElementById('modal_delete').checked = false;
                    showModalAlert('Berhasil Dihapus!', response.message, 'success');
                    setTimeout(function() {
                        location.reload();
                    }, 1500);
                },
                error: function(response) {
                    document.getElementById('modal_delete').checked = false;
                    var message = 'Terjadi kesalahan saat menghapus data.';
                    if (response.responseJSON && response.responseJSON.message) {
                        message = response.responseJSON.message;
                    }
                    showModalAlert('Gagal Menghapus!', message, 'error');
                }
            });
        });
    }
});
</script>
    <div class="container mx-auto py-6 px-4 sm:px-6 lg:px-8">
        {{-- Search & Filter Form --}}
        <div class="bg-base-100 shadow-lg rounded-lg p-6 mb-6 border border-base-200">
            <form action="{{ route('admin.karyawan') }}" method="get" enctype="multipart/form-data">
                <div class="flex flex-col md:flex-row gap-4 items-end">
                    <div class="form-control w-full md:flex-1">
                        <label class="label">
                            <span class="label-text text-base-content">Nama Karyawan Magang</span>
                        </label>
                        <input type="text" name="nama_karyawan" placeholder="Cari berdasarkan nama..."
                            class="input input-bordered w-full focus:ring-primary focus:border-primary"
                            value="{{ request()->nama_karyawan }}" />
                    </div>
                    <div class="form-control w-full md:flex-1">
                        <label class="label">
                            <span class="label-text text-base-content">Instansi</span>
                        </label>
                        <select class="select select-bordered w-full focus:ring-primary focus:border-primary"
                            name="kode_departemen">
                            <option value="">Semua Instansi</option> {{-- Opsi "Semua" --}}
                            @foreach ($departemen as $item)
                                <option value="{{ $item->kode }}" @if ($item->kode == request()->kode_departemen) selected @endif>
                                    {{ $item->nama }}</option>
                            @endforeach
                        </select>
                    </div>
                    <button type="submit"
                        class="btn btn-primary w-full md:w-auto h-full px-8 py-3 rounded-md shadow-md hover:shadow-lg transition-all duration-200 mt-4 md:mt-0">
                        <i class="ri-search-2-line text-lg"></i>
                        <span class="md:hidden">Cari</span>
                    </button>
                </div>
            </form>
        </div>

        {{-- Data Table --}}
        <div class="w-full overflow-x-auto rounded-lg shadow-lg border border-base-200 bg-base-100">
            <table class="table table-zebra w-full"> {{-- Menggunakan table-zebra untuk alternating row colors --}}
                <thead>
                    <tr>
                        <th class="text-base-content text-opacity-80">No</th>
                        <th class="text-base-content text-opacity-80">Instansi</th>
                        <th class="text-base-content text-opacity-80">Nama Lengkap</th>
                        <th class="text-base-content text-opacity-80">Foto</th>
                        <th class="text-base-content text-opacity-80">Posisi</th>
                        <th class="text-base-content text-opacity-80">Telepon</th>
                        <th class="text-base-content text-opacity-80">Email</th>
                        <th class="text-base-content text-opacity-80">Aksi</th>
                    </tr>
                </thead>
                <tbody>
                    @foreach ($karyawan as $value => $item)
                        <tr>
                            <td class="font-bold text-base-content">{{ $karyawan->firstItem() + $value }}</td>
                            <td class="text-base-content text-opacity-70">{{ $item->departemen->kode }}</td>
                            <td class="text-base-content text-opacity-70">{{ $item->nama_lengkap }}</td>
                            <td>
                                <div class="avatar">
                                    <div class="w-12 h-12 rounded-full overflow-hidden border border-base-300">
                                        @if ($item->foto)
                                            <img src="{{ asset("storage/unggah/karyawan/$item->foto") }}"
                                                alt="{{ $item->nama_lengkap }}" class="object-cover w-full h-full" />
                                        @else
                                            <img src="{{ asset('img/team-2.jpg') }}" alt="Default Profile"
                                                class="object-cover w-full h-full" />
                                        @endif
                                    </div>
                                </div>
                            </td>
                            <td class="text-base-content text-opacity-70">{{ $item->jabatan }}</td>
                            <td class="text-base-content text-opacity-70">{{ $item->telepon }}</td>
                            <td class="text-base-content text-opacity-70">{{ $item->email }}</td>
                            <td>
                                <div class="flex gap-2">
                                    <label
                                        class="btn btn-warning btn-sm btn-circle shadow-md hover:shadow-lg transition-all duration-200"
                                        for="edit_button" onclick="return edit_button('{{ $item->nik }}')">
                                        <i class="ri-pencil-fill text-lg"></i>
                                    </label>
                                    <label
                                        class="btn btn-error btn-sm btn-circle shadow-md hover:shadow-lg transition-all duration-200"
                                        onclick="return delete_button('{{ $item->nik }}', '{{ $item->nama_lengkap }}')">
                                        <i class="ri-delete-bin-line text-lg"></i>
                                    </label>
                                </div>
                            </td>
                        </tr>
                    @endforeach
                </tbody>
            </table>
            <div class="p-4 pagination">
                {{ $karyawan->links() }} {{-- Kembali ke pagination default Laravel --}}
            </div>
        </div>
    </div>

    {{-- Awal Modal Create --}}
    <input type="checkbox" id="create_modal" class="modal-toggle" />
    <div class="modal" role="dialog">
        <div class="modal-box bg-base-100 text-base-content shadow-xl border border-base-200">
            <div class="mb-4 flex items-center justify-between">
                <h3 class="text-xl font-bold">Tambah Data Karyawan Magang</h3>
                <label for="create_modal" class="btn btn-sm btn-circle btn-ghost text-base-content">
                    <i class="ri-close-large-fill text-xl"></i>
                </label>
            </div>
            <div>
                <form action="{{ route('admin.karyawan.store') }}" method="POST" enctype="multipart/form-data"
                    class="space-y-4">
                    @csrf
                    <div class="form-control">
                        <label class="label">
                            <span class="label-text font-semibold text-base-content">NIK<span
                                    class="text-error">*</span></span>
                        </label>
                        <input type="text" name="nik" placeholder="Masukkan NIK"
                            class="input input-bordered w-full focus:ring-primary focus:border-primary" required />
                        @error('nik')
                            <div class="label">
                                <span class="label-text-alt text-sm text-error">{{ $message }}</span>
                            </div>
                        @enderror
                    </div>
                    <div class="form-control">
                        <label class="label">
                            <span class="label-text font-semibold text-base-content">Instansi<span
                                    class="text-error">*</span></span>
                        </label>
                        <select name="departemen_id"
                            class="select select-bordered w-full focus:ring-primary focus:border-primary">
                            <option disabled selected>Pilih Instansi!</option>
                            @foreach ($departemen as $item)
                                <option value="{{ $item->id }}">
                                    {{ $item->nama }}</option>
                            @endforeach
                        </select>
                        @error('departemen_id')
                            <div class="label">
                                <span class="label-text-alt text-sm text-error">{{ $message }}</span>
                            </div>
                        @enderror
                    </div>
                    <div class="form-control">
                        <label class="label">
                            <span class="label-text font-semibold text-base-content">Nama Lengkap<span
                                    class="text-error">*</span></span>
                        </label>
                        <input type="text" name="nama_lengkap" placeholder="Masukkan Nama Lengkap"
                            class="input input-bordered w-full focus:ring-primary focus:border-primary" required />
                        @error('nama_lengkap')
                            <div class="label">
                                <span class="label-text-alt text-sm text-error">{{ $message }}</span>
                            </div>
                        @enderror
                    </div>
                    <div class="form-control">
                        <label class="label">
                            <span class="label-text font-semibold text-base-content">Posisi<span
                                    class="text-error">*</span></span>
                        </label>
                        <input type="text" name="jabatan" placeholder="Masukkan Posisi"
                            class="input input-bordered w-full focus:ring-primary focus:border-primary" required />
                        @error('jabatan')
                            <div class="label">
                                <span class="label-text-alt text-sm text-error">{{ $message }}</span>
                            </div>
                        @enderror
                    </div>
                    <div class="form-control">
                        <label class="label">
                            <span class="label-text font-semibold text-base-content">Telepon<span
                                    class="text-error">*</span></span>
                        </label>
                        <input type="text" name="telepon" placeholder="Masukkan Telepon"
                            class="input input-bordered w-full focus:ring-primary focus:border-primary" required />
                        @error('telepon')
                            <div class="label">
                                <span class="label-text-alt text-sm text-error">{{ $message }}</span>
                            </div>
                        @enderror
                    </div>
                    <div class="form-control">
                        <label class="label">
                            <span class="label-text font-semibold text-base-content">Email<span
                                    class="text-error">*</span></span>
                        </label>
                        <input type="email" name="email" placeholder="Masukkan Email"
                            class="input input-bordered w-full focus:ring-primary focus:border-primary" required />
                        @error('email')
                            <div class="label">
                                <span class="label-text-alt text-sm text-error">{{ $message }}</span>
                            </div>
                        @enderror
                    </div>
                    <div class="form-control">
                        <label class="label">
                            <span class="label-text font-semibold text-base-content">Password<span
                                    class="text-error">*</span></span>
                        </label>
                        <input type="password" name="password" placeholder="Masukkan Password"
                            class="input input-bordered w-full focus:ring-primary focus:border-primary" required />
                        @error('password')
                            <div class="label">
                                <span class="label-text-alt text-sm text-error">{{ $message }}</span>
                            </div>
                        @enderror
                    </div>
                    <div class="form-control">
                        <label class="label">
                            <span class="label-text font-semibold text-base-content">Foto</span>
                        </label>
                        <input type="file" name="foto" id="foto"
                            class="file-input file-input-bordered w-full file-input-primary"
                            onchange="previewImage()" />
                        @error('foto')
                            <div class="label">
                                <span class="label-text-alt text-sm text-error">{{ $message }}</span>
                            </div>
                        @enderror
                        <img class="img-preview my-3 rounded-lg shadow-sm max-h-48 object-cover object-center"
                            style="display: none;" /> {{-- Default hidden --}}
                    </div>
                    <div class="modal-action mt-6">
                        <button type="submit"
                            class="btn btn-primary w-full text-white shadow-md hover:shadow-lg transition-all duration-200">Simpan
                            Data</button>
                    </div>
                </form>
            </div>
        </div>
    </div>
    {{-- Akhir Modal Create --}}

    {{-- Awal Modal Edit --}}
    <input type="checkbox" id="edit_button" class="modal-toggle" />
    <div class="modal" role="dialog">
        <div class="modal-box bg-base-100 text-base-content shadow-xl border border-base-200">
            <div class="mb-4 flex items-center justify-between">
                <h3 class="text-xl font-bold">Ubah Data Karyawan Magang</h3>
                <label for="edit_button" class="btn btn-sm btn-circle btn-ghost text-base-content">
                    <i class="ri-close-large-fill text-xl"></i>
                </label>
            </div>
            <div>
                <form action="{{ route('admin.karyawan.update') }}" method="POST" enctype="multipart/form-data"
                    class="space-y-4">
                    @csrf
                    <input type="hidden" name="nik_lama">
                    <div class="form-control">
                        <label class="label">
                            <span class="label-text font-semibold text-base-content">NIK<span
                                    class="text-error">*</span></span>
                            <span class="label-text-alt" id="loading_edit1"></span>
                        </label>
                        <input type="text" name="nik" placeholder="Masukkan NIK"
                            class="input input-bordered w-full focus:ring-primary focus:border-primary" required />
                        @error('nik')
                            <div class="label">
                                <span class="label-text-alt text-sm text-error">{{ $message }}</span>
                            </div>
                        @enderror
                    </div>
                    <div class="form-control">
                        <label class="label">
                            <span class="label-text font-semibold text-base-content">Instansi<span
                                    class="text-error">*</span></span>
                            <span class="label-text-alt" id="loading_edit2"></span>
                        </label>
                        <select name="departemen_id" id='departemen_id'
                            class="select select-bordered w-full focus:ring-primary focus:border-primary">
                        </select>
                        @error('departemen_id')
                            <div class="label">
                                <span class="label-text-alt text-sm text-error">{{ $message }}</span>
                            </div>
                        @enderror
                    </div>
                    <div class="form-control">
                        <label class="label">
                            <span class="label-text font-semibold text-base-content">Nama Lengkap<span
                                    class="text-error">*</span></span>
                            <span class="label-text-alt" id="loading_edit3"></span>
                        </label>
                        <input type="text" name="nama_lengkap" placeholder="Masukkan Nama Lengkap"
                            class="input input-bordered w-full focus:ring-primary focus:border-primary" required />
                        @error('nama_lengkap')
                            <div class="label">
                                <span class="label-text-alt text-sm text-error">{{ $message }}</span>
                            </div>
                        @enderror
                    </div>
                    <div class="form-control">
                        <label class="label">
                            <span class="label-text font-semibold text-base-content">Posisi<span
                                    class="text-error">*</span></span>
                            <span class="label-text-alt" id="loading_edit4"></span>
                        </label>
                        <input type="text" name="jabatan" placeholder="Masukkan Posisi"
                            class="input input-bordered w-full focus:ring-primary focus:border-primary" required />
                        @error('jabatan')
                            <div class="label">
                                <span class="label-text-alt text-sm text-error">{{ $message }}</span>
                            </div>
                        @enderror
                    </div>
                    <div class="form-control">
                        <label class="label">
                            <span class="label-text font-semibold text-base-content">Telepon<span
                                    class="text-error">*</span></span>
                            <span class="label-text-alt" id="loading_edit5"></span>
                        </label>
                        <input type="text" name="telepon" placeholder="Masukkan Telepon"
                            class="input input-bordered w-full focus:ring-primary focus:border-primary" required />
                        @error('telepon')
                            <div class="label">
                                <span class="label-text-alt text-sm text-error">{{ $message }}</span>
                            </div>
                        @enderror
                    </div>
                    <div class="form-control">
                        <label class="label">
                            <span class="label-text font-semibold text-base-content">Email<span
                                    class="text-error">*</span></span>
                            <span class="label-text-alt" id="loading_edit6"></span>
                        </label>
                        <input type="email" name="email" placeholder="Masukkan Email"
                            class="input input-bordered w-full focus:ring-primary focus:border-primary" required />
                        @error('email')
                            <div class="label">
                                <span class="label-text-alt text-sm text-error">{{ $message }}</span>
                            </div>
                        @enderror
                    </div>
                    <div class="form-control">
                        <label class="label">
                            <span class="label-text font-semibold text-base-content">Foto</span>
                            <span class="label-text-alt" id="loading_edit7"></span>
                        </label>
                        <input type="file" name="foto" id="foto_edit"
                            class="file-input file-input-bordered w-full file-input-warning"
                            onchange="previewImageEdit()" />
                        @error('foto')
                            <div class="label">
                                <span class="label-text-alt text-sm text-error">{{ $message }}</span>
                            </div>
                        @enderror
                        <img class="foto-edit-preview my-3 rounded-lg shadow-sm max-h-48 object-cover object-center"
                            style="display: none;" /> {{-- Default hidden --}}
                    </div>
                    <div class="modal-action mt-6">
                        <button type="submit"
                            class="btn btn-warning w-full text-white shadow-md hover:shadow-lg transition-all duration-200">Perbarui
                            Data</button>
                    </div>
                </form>
            </div>
        </div>
    </div>
    {{-- Akhir Modal Edit --}}

    {{-- Global Modal Notifikasi DaisyUI --}}
    <input type="checkbox" id="modal_alert" class="modal-toggle" />
    <div class="modal" role="dialog">
        <div class="modal-box bg-base-100 text-base-content border border-base-200 shadow-xl">
            <h3 id="modal_title" class="font-bold text-lg">Judul</h3>
            <p id="modal_message" class="py-4 text-base-content text-opacity-80"></p>
            <div class="modal-action">
                <label for="modal_alert" class="btn btn-primary">OK</label>
            </div>
        </div>
    </div>

    {{-- Modal Konfirmasi Hapus --}}
    <input type="checkbox" id="modal_delete" class="modal-toggle" />
    <div class="modal" role="dialog">
        <div class="modal-box bg-base-100 text-base-content border border-base-200 shadow-xl">
            <h3 class="font-bold text-lg text-error">Konfirmasi Hapus Data</h3>
            <p id="delete_message" class="py-4"></p>
            <div class="modal-action">
                <button id="confirm_delete_btn" class="btn btn-error">Ya, Hapus</button>
                <label for="modal_delete" class="btn">Batal</label>
            </div>
        </div>
    </div>

    {{-- Awal Modal Import --}}
    <input type="checkbox" id="import_modal" class="modal-toggle" />
    <div class="modal" role="dialog">
        <div class="modal-box bg-base-100 text-base-content shadow-xl border border-base-200">
            <div class="mb-4 flex items-center justify-between">
                <h3 class="text-xl font-bold">Import Data Karyawan Magang</h3>
                <label for="import_modal" class="btn btn-sm btn-circle btn-ghost text-base-content">
                    <i class="ri-close-large-fill text-xl"></i>
                </label>
            </div>
            <div>
                <form action="{{ route('admin.karyawan.import') }}" method="POST" enctype="multipart/form-data"
                    class="space-y-4">
                    @csrf
                    <div class="form-control">
                        <label class="label">
                            <span class="label-text font-semibold text-base-content">Pilih File Excel/CSV<span
                                    class="text-error">*</span></span>
                        </label>
                        <input type="file" name="file"
                            class="file-input file-input-bordered w-full file-input-info" required />
                        @error('file')
                            <div class="label">
                                <span class="label-text-alt text-sm text-error">{{ $message }}</span>
                            </div>
                        @enderror
                        <p class="text-sm text-base-content text-opacity-70 mt-2">
                            Pastikan format file Anda memiliki header berikut: NIK, ID Departemen (atau Nama Instansi),
                            Nama Lengkap, Posisi, Telepon, Email, Password (opsional).
                            Untuk update data, NIK yang sama akan menimpa data yang ada. Password wajib diisi untuk
                            karyawan baru.
                        </p>
                    </div>
                    <div class="modal-action mt-6">
                        <button type="submit"
                            class="btn btn-info w-full text-white shadow-md hover:shadow-lg transition-all duration-200">Import
                            Data</button>
                    </div>
                </form>
            </div>
        </div>
    </div>
    {{-- Akhir Modal Import --}}


  
</x-app-layout>
