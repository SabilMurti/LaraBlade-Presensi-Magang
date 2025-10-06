<?php

namespace App\Http\Controllers;

use App\Models\Departemen;
use App\Models\Karyawan;
use Illuminate\Http\Request;
use Carbon\Carbon;
use Illuminate\Support\Facades\Hash;
use Illuminate\Support\Facades\Storage;
use Illuminate\Validation\Rule;


// Tambahkan use statement untuk PhpSpreadsheet
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use Illuminate\Support\Facades\DB; // Untuk transaksi database

use App\Http\Traits\ExcelExportTrait;
class KaryawanController extends Controller
{

    use ExcelExportTrait;


    // public function indexAdmin(Request $request)
    // {
    //     $title = "Data Karyawan Magang"; // Perbarui judul sesuai permintaan sebelumnya

    //     $departemen = Departemen::get();

    //     $query = Karyawan::join('departemen as d', 'karyawan.departemen_id', '=', 'd.id')
    //         ->select('karyawan.nik', 'd.nama as nama_departemen', 'karyawan.nama_lengkap', 'karyawan.foto', 'karyawan.jabatan', 'karyawan.telepon', 'karyawan.email', 'karyawan.departemen_id') // Tambah departemen_id untuk keperluan export
    //         ->orderBy('d.kode', 'asc')
    //         ->orderBy('karyawan.nama_lengkap', 'asc');

    //     if ($request->nama_karyawan) {
    //         $query->where('karyawan.nama_lengkap', 'like', '%' . $request->nama_karyawan . '%');
    //     }
    //     if ($request->kode_departemen) {
    //         $query->where('d.kode', 'like', '%' . $request->kode_departemen . '%');
    //     }
    //     $karyawan = $query->paginate(10);

    //     return view('admin.karyawan.index', compact('title', 'karyawan', 'departemen'));
    // }

    public function exportKaryawan()
    {
        $fileName = 'data_karyawan_' . Carbon::now()->format('Ymd_His');

        $headings = [
            'NIK',
            'ID Departemen', // Untuk referensi
            'Nama Instansi', // Jika ingin ada nama instansi di export
            'Nama Lengkap',
            'Posisi',
            'Telepon',
            'Email'
        ];

        // Ambil data dari database
        $karyawanData = Karyawan::join('departemen as d', 'karyawan.departemen_id', '=', 'd.id')
            ->select(
                'karyawan.nik',
                'karyawan.departemen_id', // ID departemen penting untuk import
                'd.nama as nama_departemen', // Nama departemen untuk informasi
                'karyawan.nama_lengkap',
                'karyawan.jabatan', // 'jabatan' dari DB akan jadi 'Posisi' di Excel
                'karyawan.telepon',
                'karyawan.email'
            )
            ->get();

        // Siapkan data dalam format array untuk diexport
        $exportData = $karyawanData->map(function ($item) {
            return [
                $item->nik,
                $item->departemen_id,
                $item->nama_departemen,
                $item->nama_lengkap,
                $item->jabatan,
                $item->telepon,
                $item->email
            ];
        });

        return $this->exportDataToExcel($fileName, $headings, $exportData, 'xlsx'); // Bisa juga 'csv'
    }

   public function importKaryawan(Request $request)
{
    $request->validate([
        'file' => 'required|mimes:xlsx,xls,csv',
    ], [
        'file.required' => 'Harap pilih file untuk diunggah.',
        'file.mimes' => 'File harus berformat Excel (xlsx, xls) atau CSV.',
    ]);

    $file = $request->file('file');
    $filePath = $file->getRealPath();
    $spreadsheet = IOFactory::load($filePath);
    $sheetData = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);

    // Asumsi baris pertama adalah header
    $header = array_map('trim', $sheetData[1]);
    unset($sheetData[1]);

    $expectedHeaders = [
        'NIK',
        'ID Departemen',
        'Nama Lengkap',
        'Posisi',
        'Telepon',
        'Email',
        'Password'
    ];

    // Validasi header
    foreach ($expectedHeaders as $expected) {
        if ($expected === 'Password') {
            continue; // Password opsional
        }
        if (!in_array($expected, $header)) {
            return to_route('admin.karyawan')->with('error', "Kolom '$expected' tidak ditemukan di file Excel/CSV. Harap periksa template.");
        }
    }

    $importedRows = 0; // Ubah dari array ke counter
    $errors = [];

    DB::beginTransaction();
    try {
        foreach ($sheetData as $rowNum => $row) {
            // Lewati baris kosong
            if (empty(array_filter($row))) {
                continue;
            }

            // Mapping kolom berdasarkan header
            $data = [];
            foreach ($header as $colLetter => $colName) {
                $data[$colName] = $row[$colLetter] ?? null;
            }

            // Data yang akan diimpor
            $nik = trim($data['NIK'] ?? '');
            $departemenIdentifier = trim($data['ID Departemen'] ?? ($data['Nama Instansi'] ?? ''));
            $namaLengkap = trim($data['Nama Lengkap'] ?? '');
            $jabatan = trim($data['Posisi'] ?? '');
            $telepon = trim($data['Telepon'] ?? '');
            $email = trim($data['Email'] ?? '');
            $password = isset($data['Password']) ? trim($data['Password']) : null;

            // Validasi data
            $rules = [
                'nik' => 'required|string|max:255',
                'departemen_identifier' => 'required',
                'nama_lengkap' => 'required|string|max:255',
                'jabatan' => 'required|string|max:255',
                'telepon' => 'required|string|max:15',
                'email' => 'required|string|email|max:255',
                'password' => 'nullable|string|min:6',
            ];

            $validator = \Illuminate\Support\Facades\Validator::make([
                'nik' => $nik,
                'departemen_identifier' => $departemenIdentifier,
                'nama_lengkap' => $namaLengkap,
                'jabatan' => $jabatan,
                'telepon' => $telepon,
                'email' => $email,
                'password' => $password,
            ], $rules);

            if ($validator->fails()) {
                $errors[] = "Baris $rowNum: " . implode(", ", $validator->errors()->all());
                continue;
            }

            // Temukan ID Departemen
            $departemen = null;
            if (is_numeric($departemenIdentifier)) {
                $departemen = Departemen::find($departemenIdentifier);
            } else {
                $departemen = Departemen::where('nama', $departemenIdentifier)->first();
            }

            if (!$departemen) {
                $errors[] = "Baris $rowNum: Instansi '$departemenIdentifier' tidak ditemukan.";
                continue;
            }
            $departemenId = $departemen->id;

            // Cek NIK dan Email unik
            $existingKaryawan = Karyawan::where('nik', $nik)->first();
            if ($existingKaryawan && $existingKaryawan->email !== $email) {
                $errors[] = "Baris $rowNum: NIK '$nik' sudah ada dengan email yang berbeda.";
                continue;
            }
            
            $existingEmailKaryawan = Karyawan::where('email', $email)->first();
            if ($existingEmailKaryawan && $existingEmailKaryawan->nik !== $nik) {
                $errors[] = "Baris $rowNum: Email '$email' sudah digunakan oleh karyawan dengan NIK berbeda.";
                continue;
            }

            $karyawanData = [
                'departemen_id' => $departemenId,
                'nama_lengkap' => $namaLengkap,
                'jabatan' => $jabatan,
                'telepon' => $telepon,
                'email' => $email,
            ];

            if ($existingKaryawan) {
                // Update karyawan yang sudah ada
                if ($password) {
                    $karyawanData['password'] = Hash::make($password);
                }
                Karyawan::where('nik', $nik)->update($karyawanData);
                $importedRows++;
            } else {
                // Buat karyawan baru - password wajib
                if (!$password) {
                    $errors[] = "Baris $rowNum: Password wajib diisi untuk karyawan baru dengan NIK '$nik'.";
                    continue;
                }
                $karyawanData['password'] = Hash::make($password);
                Karyawan::create(array_merge(['nik' => $nik], $karyawanData));
                $importedRows++;
            }
        }

        // PENTING: Cek apakah ada error ATAU tidak ada data yang berhasil diimport
        if (!empty($errors)) {
            DB::rollBack();
            $errorMessage = "Gagal mengimpor data. Beberapa baris memiliki kesalahan:<br>";
            foreach ($errors as $error) {
                $errorMessage .= "- $error<br>";
            }
            return to_route('admin.karyawan')->with('error', $errorMessage);
        }

        // Cek apakah ada data yang berhasil diimport
        if ($importedRows === 0) {
            DB::rollBack();
            return to_route('admin.karyawan')->with('error', 'Tidak ada data yang berhasil diimpor. Periksa format file Anda.');
        }

        DB::commit();
        return to_route('admin.karyawan')->with('success', "Data Karyawan berhasil diimpor ($importedRows baris).");
        
    } catch (\Exception $e) {
        DB::rollBack();
        return to_route('admin.karyawan')->with('error', 'Gagal mengimpor data: ' . $e->getMessage());
    }
}

    public function index()
    {
        $title = "Profile";
        $karyawan = Karyawan::where('nik', auth()->guard('karyawan')->user()->nik)->first();
        return view('dashboard.profile.index', compact('title', 'karyawan'));
    }

    public function update(Request $request)
    {
        $karyawan = Karyawan::where('nik', auth()->guard('karyawan')->user()->nik)->first();

        if ($request->hasFile('foto')) {
            $foto = $karyawan->nik . "." . $request->file('foto')->getClientOriginalExtension();
        } else {
            $foto = $karyawan->foto;
        }

        if ($request->password != null) {
            $update = Karyawan::where('nik', auth()->guard('karyawan')->user()->nik)->update([
                'nama_lengkap' => $request->nama_lengkap,
                'telepon' => $request->telepon,
                'password' => Hash::make($request->password),
                'foto' => $foto,
                'updated_at' => Carbon::now(),
            ]);
        } elseif ($request->password == null) {
            $update = Karyawan::where('nik', auth()->guard('karyawan')->user()->nik)->update([
                'nama_lengkap' => $request->nama_lengkap,
                'telepon' => $request->telepon,
                'foto' => $foto,
                'updated_at' => Carbon::now(),
            ]);
        }

        if ($update) {
            if ($request->hasFile('foto')) {
                $folderPath = "public/unggah/karyawan/";
                $request->file('foto')->storeAs($folderPath, $foto);
            }
            return redirect()->back()->with('success', 'Profile updated successfully');
        } else {
            return redirect()->back()->with('error', 'Profile updated failed');
        }
    }

    public function indexAdmin(Request $request)
    { {
            $title = "Data Karyawan Magang"; // Perbarui judul sesuai permintaan sebelumnya

            $departemen = Departemen::get();

            $query = Karyawan::join('departemen as d', 'karyawan.departemen_id', '=', 'd.id')
                ->select('karyawan.nik', 'd.nama as nama_departemen', 'karyawan.nama_lengkap', 'karyawan.foto', 'karyawan.jabatan', 'karyawan.telepon', 'karyawan.email', 'karyawan.departemen_id') // Tambah departemen_id untuk keperluan export
                ->orderBy('d.kode', 'asc')
                ->orderBy('karyawan.nama_lengkap', 'asc');

            if ($request->nama_karyawan) {
                $query->where('karyawan.nama_lengkap', 'like', '%' . $request->nama_karyawan . '%');
            }
            if ($request->kode_departemen) {
                $query->where('d.kode', 'like', '%' . $request->kode_departemen . '%');
            }
            $karyawan = $query->paginate(10);

            return view('admin.karyawan.index', compact('title', 'karyawan', 'departemen'));
        }
    }

    public function store(Request $request)
    {
        $data = $request->validate([
            'nik' => 'required|unique:karyawan,nik',
            'departemen_id' => 'required',
            'nama_lengkap' => 'required|string|max:255',
            'foto' => 'nullable|image|mimes:jpeg,png,jpg|max:2048',
            'jabatan' => 'required|string|max:255',
            'telepon' => 'required|string|max:15',
            'email' => 'required|string|email|max:255|unique:karyawan,email',
            'password' => 'required',
        ]);
        $data['password'] = Hash::make($data['password']);
        if ($request->hasFile('foto')) {
            $foto = $request->nik . "." . $request->file('foto')->getClientOriginalExtension();
        }

        $create = Karyawan::create($data);

        if ($create) {
            if ($request->hasFile('foto')) {
                $folderPath = "public/unggah/karyawan/";
                $request->file('foto')->storeAs($folderPath, $foto);
            }
            return to_route('admin.karyawan')->with('success', 'Data Karyawan berhasil disimpan');
        } else {
            return to_route('admin.karyawan')->with('error', 'Data Karyawan gagal disimpan');
        }
    }

    public function edit(Request $request)
    {
        $data = Karyawan::where('nik', $request->nik)->first();
        return $data;
    }

    public function updateAdmin(Request $request)
    {
        $karyawan = Karyawan::where('nik', $request->nik_lama)->first();
        $data = $request->validate([
            'nik' => ['required', Rule::unique('karyawan')->ignore($karyawan)],
            'departemen_id' => 'required',
            'nama_lengkap' => 'required|string|max:255',
            'foto' => 'nullable|image|mimes:jpeg,png,jpg|max:2048',
            'jabatan' => 'required|string|max:255',
            'telepon' => 'required|string|max:15',
            'email' => ['required', 'email', Rule::unique('karyawan')->ignore($karyawan)],
        ]);
        if ($request->hasFile('foto')) {
            $data['foto'] = $request->nik . "." . $request->file('foto')->getClientOriginalExtension();
        }

        $update = Karyawan::where('nik', $request->nik_lama)->update($data);

        if ($update) {
            if ($request->hasFile('foto')) {
                $folderPath = "public/unggah/karyawan/";
                $request->file('foto')->storeAs($folderPath, $data['foto']);
            }
            return to_route('admin.karyawan')->with('success', 'Data Karyawan berhasil diperbarui');
        } else {
            return to_route('admin.karyawan')->with('error', 'Data Karyawan gagal diperbarui');
        }
    }

    public function delete(Request $request)
    {
        try {
            $data = Karyawan::where('nik', $request->nik)->firstOrFail();

            // Simpan nama file sebelum delete
            $foto = $data->foto;

            // Coba hapus data
            $delete = $data->delete();

            // Jika berhasil dan ada foto, hapus dari storage
            if ($delete && $foto) {
                $folderPath = "public/unggah/karyawan/";
                Storage::delete($folderPath . $foto);
            }

            return response()->json([
                'success' => true,
                'message' => 'Data karyawan berhasil dihapus.'
            ]);
        } catch (\Illuminate\Database\QueryException $e) {
            if ($e->getCode() == 23000) {
                // Ini adalah error karena foreign key constraint
                return response()->json([
                    'success' => false,
                    'message' => 'Tidak dapat menghapus karyawan ini karena masih memiliki data pengajuan presensi yang terhubung.'
                ], 400);
            }

            // Error SQL lain
            return response()->json([
                'success' => false,
                'message' => 'Terjadi kesalahan pada database: ' . $e->getMessage(),
            ], 500);
        } catch (\Exception $e) {
            // Error umum lainnya
            return response()->json([
                'success' => false,
                'message' => 'Terjadi kesalahan: ' . $e->getMessage(),
            ], 500);
        }
    }
}
