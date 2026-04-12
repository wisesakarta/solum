# Learning Wiki: Saka Note

Selamat datang di jurnal pembelajaran Saka Note! Di sini kita mencatat evolusi arsitektur dan kebijaksanaan rekayasa yang kita temukan.

## 🏙️ Arsitektur (Peta Kota)

Saka Note dibangun dengan struktur **Single Activity, Multi-Composable**.
- **MainActivity**: Stasiun pusat yang memulai segalanya.
- **EditorApp.kt**: Jantung dari aplikasi, tempat di mana `TechnicalStandardNoteMobileApp` mengatur orkestrasi editor, tab, dan sistem keamanan.
- **Root Box Pattern**: Kita menggunakan satu `Box` besar sebagai root untuk memungkinkan layering (seperti overlay App Lock dan feedback loading) di atas UI utama.
- **Recursive Tree**: Struktur folder menggunakan pola *Adjacency List* di database (parentId) yang dirender secara rekursif di UI menggunakan `RecursiveFolderList`. Analogi: "Aplikasi ini seperti akar pohon yang terus bercabang ke bawah tanpa batas."

## 🛠️ Keputusan Teknis (Alasan Mengapa)

- **Jetpack Compose**: Kita memilih pendekatan deklaratif karena memudahkan manajemen state yang kompleks seperti sinkronisasi tab dan App Lock.
- **Custom Font Loading**: Kita mendukung font eksternal (.ttf/.otf) untuk memberikan kebebasan tipografi kepada pengguna, yang kita simpan secara aman menggunakan Android Scoped Storage.
- **UUID Identifiers**: Kita beralih dari ID bertipe `Long` ke `String` (UUID) untuk memastikan integritas data yang lebih baik saat melakukan operasi sinkronisasi dan identifikasi unik di struktur folder yang dalam.
- **Adjacency List + CTE**: Kita menggunakan SQLite *Recursive CTE* (`WITH RECURSIVE`) untuk mengambil seluruh hierarki folder dalam satu query tunggal.

## 🐛 Cerita Perang (Bug & Perbaikan)

### Log: 2026-04-12 - Kegelapan Mutlak dari `nullptr`
- **Masalah**: Panel *Custom Scrollbar* dikanan Editor memuntahkan warna hitam solid yang pekat seperti lubang hitam (*Black Hole*) saat menggunakan tema *Studio Paper* (Light Mode).
- **Penyebab**: Proses pembersihan kanvas (*clear frame*) pada rutin `WM_PAINT` D2D di *scrollbar* menggunakan instruksi `context->Clear(nullptr)`. D2D menginterpretasikan memori kosong (*nullptr*) sebagai spektrum warna ARGB `(0,0,0,0)`. Karena jendela anakan (*Child Window*) di Win32 tidak secara otomatis di-komposit dengan kapabilitas lapisan tembus pandang (*WS_EX_LAYERED* / Alpha Blend) tanpa konfigurasi DWM yang mahal, warna transparansi 0 alpha ini diterjemahkan oleh OS sebagai spektrum RGB `0x000000` (Hitam Mutlak) yang kedap air.
- **Pelajaran**: "Tidak ada transparansi gratis di Desktop Win32 kuno". Saat menggunakan *HwndRenderTarget* yang tak transparan/opaque, kanvas tidak boleh dibersihkan dengan `nullptr`. Saya mengarahkan instruksi pelitur *background* agar mencomot parameter warna latar belakang aktual yang sedang aktif di sistem (`DesignSystem::Color::kLightBg` atau DarkBg) secara *Real-Time*, lalu membilas kanvas sepenuhnya agar panel membaur sempurna dengan halaman Editor layaknya bunglon.

### Log: 2026-04-12 - Presisi Koordinat Absolut vs Konstanta Kasar
- **Masalah**: Panel Command Palette terlihat tumpang-tindih (*overlapping*) secara tidak rapi dengan garis pemisah (*separator line*) antara tab dan editor. Parameter Y aslinya adalah konstanta kasar `80 * scale` dari atap jendela.
- **Penyebab**: Tinggi panel *Title Bar* bergantung secara variabel pada versi OS Windows (10 vs 11 vs konfigurasi DPI). Menebak `80` piksel dari atap akan selalu meleset dari tepi kanvas editor.
- **Pelajaran**: "Never hardcode window offset coordinates." Jangan menebak piksel. Karena Command Palette harus muncul sejajar dengan garis tabs, kita menangkap nilai kordinat `rcEditor.top` langsung dari elemen editor absolut di layar (*Global Physical Screen Coordinates*), melenyapkan cacat penempatan jendela secara permanen ke tingkat piksel *(pixel-perfect alignment)*.

### Log: 2026-04-12 - Resolusi "Double-Scaling" Direct2D & Win32
- **Masalah**: Pada monitor beresolusi tinggi (High-DPI 125% atau 150%), jendela antarmuka palet (Command Palette) gagal memuat semua barisn teks, terpotong persis di bagian bawah secara proporsional dengan rasio skala monitor layar (*UI Clipping*).
- **Penyebab**: Aplikasi ini dikonfigurasi sebagai *Per-Monitor DPI Aware V2*. Elemen rendering *Direct2D* (beserta *DirectWrite*) mengadopsi DPI monitor langsung dan mengamplifikasi besaran ukuran pixel teks di internal secara otonom (Logical Pixel multiplier). Namun, dimensi wadah kaca pembungkus form *Win32* tidak dikalikan dan dioper mentah-mentah menggunakan konstanta fisik, akibatnya elemen antarmuka yang membengkak ditabrakkan dengan batasan form yang terlalu sempit.
- **Pelajaran**: Dalam lingkungan antarmuka gabungan Win32 & D2D, jika wadah menggunakan fungsi OS natif seperti `SetWindowPos`, dimensi fisik wajib dikalibrasi terlebih dahulu secara manual dengan skalar `GetDeviceCaps(hdc, LOGPIXELSY) / 96.0f` milik monitor terkini (skalar fisik), lalu barulah injeksikan murni variabel *Logical Pixels* 100% (seperti besaran font 15.0f atau padding 20pt) kepada D2D.

### Log: 2026-04-12 - Eksorsisme DWrite Private Font Cache
- **Masalah**: Font kustom *Akkurat Mono LL* selalu ditolak saat dirender oleh `Command Palette`, memaksa sistem jatuh kembali pada "AI Slop fallback" *Segoe UI* proporsional.
- **Penyebab**: Fungsi pemuatan font aplikasi Win32 (`AddFontResourceExW`) memblokir visibilitas font dari OS menggunakan argumen rahasia sesi `FR_PRIVATE`. DWrite, yang berada di *pipeline* rendering berbeda, terlahir dengan anatomi "buta" terhadap instalasi font privat aplikasi.  
- **Pelajaran**: "Perbaikan satu baris untuk masalah seribu bahasa." Dari pada menulis *Factory Object Font-Loader C++* ratusan baris, saya membuang flag `FR_PRIVATE` dan mengubahnya menjadi argument `0`. Ini menginstruksikan sistem Windows untuk menyekresi *Akkurat* secara sementara & global ke dalam tabel matriks font komputer. DWrite akhirnya mampu menemukan font tersebut seketika setelah *Initialize Session*.

### Log: 2026-04-12 - Ilusi Kontrol Teks di Atas Kanvas D2D
- **Masalah**: Form isian teks pada fitur pencarian Command Palette hanya bisa mendeteksi ketikan standar (`A..Z`), kehilangan integrasi dasar esensial komputer (seleksi mouse, panah navigasi, *Caret* ketik, bahkan manipulasi mutlak `Ctrl+A` / `V`).
- **Penyebab**: Kotak teks tersebut ternyata hanyalah kanvas grafis palsu. Ketikan di tangkap secara mentah melalui hook `WM_CHAR` dan `std::wstring` hanya dipompa ke layar tanpa mesin teks *engine*.
- **Pelajaran**: Gunakan *Native Toolkit* untuk penyelesaian elegan. Saya mensintetis kontrol `EDIT` bawaan Windows (Subclass Win32) dan menggabungkannya tepat secara fisik (*Overlapping Invisible Overlay*) di atas kordinat form *D2D*. D2D hanya diam melukis estetika latar desain *(Emil Background)*, sementara `EDIT` secara sekejap mengemulasikan kompleksitas kursor pengetikan asli komputer di atasnya secara sinergis. C++ Subclassing memberikan pengarahan lalu-lintas pintasan keyboard seutuhnya kembali ke *Palette* bila tak terkait mengetik.

### Log: 2026-04-11 - Kasus "Brace yang Hilang"
- **Masalah**: Build gagal total pada tahap `kaptGenerateStubs` dengan pesan error `Expecting '}'` di baris terakhir file (2999) dan `Unresolved Reference` untuk `zIndex`.
- **Penyebab**: Ternyata ada satu brace `}` yang lupa ditutup pada `Column` utama di tengah file, dan import `zIndex` berada di package yang salah.
- **Pelajaran**: Kapt sangat sensitif. Jika build gagal di baris terakhir file, periksa balance brace (`{}`) di seluruh file, terutama jika baru saja melakukan refactor besar pada fungsi root.

### Log: 2026-04-11 - Stabilisasi Build & Memori
- **Masalah**: JVM Crash (`hs_err_pid`) dan build gagal dengan `Unresolved reference: Folder` setelah refactor besar.
- **Penyebab**: Mesin host kehabisan memori (Paging File Exhausted) dan `Folder` icon ternyata berada di modul `material-icons-extended` yang belum diimpor (M3 BOM tidak memasukkannya secara default).
- **Pelajaran**: (1) Saat memori kritis, limit eksekusi Gradle (`parallel=false`, `workers=1`). (2) Jangan asumsikan semua icon Material tersedia di *core set*; icon kategori `Rounded` sering kali hanya ada di *extended set*.

### Log: 2026-04-11 - Resolusi Import Icon Material
- **Masalah**: Kegagalan build akibat referensi `unresolved` untuk `MoreVert` dan `Settings`.
- **Penyebab**: Transisi UI yang menggunakan icon dari set `Rounded` memerlukan import eksplisit yang belum terdaftar di bagian atas file `EditorApp.kt`.
### Log: 2026-04-11 - Keunggulan Matrix Skew vs RotationX
- **Masalah**: `RotationX` (perspektif) menghasilkan bentuk trapesium yang tidak cocok dengan estetika "industrial geometric" pada folder React Bits.
- **Penyebab**: Transformasi perspektif standar Compose mensimulasikan kamera, sedangkan desain aslinya menggunakan transformasi linear shear (skew).
- **Pelajaran**: Untuk desain skeuomorphic yang sangat geometris, gunakan `androidx.compose.ui.graphics.Matrix` dan `drawContext.canvas.concat()`. Ini memungkinkan kontrol tingkat rendah pada matriks transformasi (seperti Skew X) yang tidak tersedia di `graphicsLayer` standar.

### Log: 2026-04-11 - Resolusi "Chaotic Matrix Smear" (Rendering Glitch)
- **Masalah**: Muncul artefak hitam bergerigi (*jagged shapes*) yang memotong layar saat folder diklik.
- **Penyebab**: Penggunaan matriks manual 4x4 tanpa kompensasi origin yang tepat menyebabkan koordinat drawing "meluber" (*state leak*) keluar batas layar.
- **Pelajaran**: Selalu gunakan primitif Canvas (`save/restore`, `translate`, `skew`, `scale`) dengan **Pivot Point** yang jelas (misal: bottom-center). Menggeser origin ke titik jangkar sebelum transformasi jauh lebih aman dan stabil daripada manipulasi matriks manual yang kompleks.

### Log: 2026-04-11 - Rahasia Anti-Aliasing: CompositingStrategy.Offscreen
- **Masalah**: Tepian komponen yang dirotasi secara 3D (`rotationX`, `rotateY`) terlihat bergerigi (*pixelated*) dan tidak premium.
- **Penyebab**: Rendering standar Compose terkadang mematikan antialiasing pada buffer transisi untuk meningkatkan performa.
- **Pelajaran**: Gunakan `compositingStrategy = CompositingStrategy.Offscreen` di dalam `graphicsLayer` untuk memaksa rendering ke buffer terpisah yang mendukung antialiasing penuh. Ini adalah "kunci" untuk mendapatkan tampilan 3D yang sangat halus (premium) seperti pada referensi UI tingkat tinggi.

### Log: 2026-04-11 - Resolusi Overlap (Emil Kowalski's Principles)
- **Masalah**: Efek *skew* pada flap folder menyebabkan komponen meluas (*flare*) keluar dari *bounding box* aslinya, sehingga menutupi atau berbenturan dengan folder di sebelahnya pada list (*menjulang terlalu lebar*).
- **Penyebab**: CSS affine transform `skew` memindahkan koordinat pixel secara fisik ke luar lebar container.
- **Pelajaran**: Mengikuti prinsip *Design Engineering* Emil Kowalski, kita **JANGAN PERNAH menganimasikan properti layout (padding/margin/width)** untuk mendorong komponen lain. Solusi absolut: Perkecil/susutkan skala ukuran komponen (*Scalling Absorbance*) melalui komposit `Modifier.scale()` sebesar nilai pelebaran *flare*, sehingga *flare* terekspansi memakan jarak susut tanpa mengubah jejak area layout grid native!

### Log: 2026-04-12 - Mathematical Tiering / Cadence Presisi (AnimeJS Staggering)
- **Masalah**: Animasi kertas dalam map terlihat "*menjulang terlalu tinggi menembus atap map*" padahal secara render sebenarnya berada di dalam kanvas. Celah tepian map depan terasa sesak menimpa kertas.
- **Penyebab**: Dimensi lebar diatur serentak `90%`, `80%`, `70%`, dan kordinasi pergeseran Y-axis nya terlalu drastis menjauh dari _Base Line_.
- **Pelajaran**: Pinjam prinsip *Stagger* Anime.js—animasi tumpukan (*stack*) harus mengikuti hirarki cadens (ritme) yang ketat. 
  1. Perkecil ruang komposit kordinat lebar ke kurva tangga eksak: `60%`, `70%`, `80%` memberikan *margin bantalan* sisi (hingga 20% kosong) yang luas dan eksklusif. 
  2. Kalkulasi batas absolut atap map `(y=0.dp)` dan atap _overlap_ `(y=48.dp)`. Posisikan rentang offset di parameter interval presisi 8dp `(y=16/24/32)`. Bukannya membesar-besarkan efek agar *lebay* menarik perhatian, cukup 8dp selisih di dimensi Y Axis sudah menstimulasi kesan kedalaman optikal yang luar biasa bersih tanpa mencuat menerobos tepi map tab di punggungnya.

### Log: 2026-04-12 - Eksorsisme Win32 Legacy Rendering
- **Masalah**: Rendering *tab spin chrome* dan *premium header* pada aplikasi Desktop C++ masih terlihat seperti Windows 95 (garis batas 1px kasar dan teks putih yang tidak nampak pada tema *Light*).
- **Penyebab**: Penggunaan instruksi melukis kuno dari pustaka Win32 klasik (`Polygon` bergerigi) ditambah nilai desimal warna (*hardcode*) yang kebal terhadap transisi pergantian sistem tema antarmuka "Studio Paper" / "Studio Charcoal".
- **Pelajaran**: 
  1. **Tokenization in GDI/Direct2D**: Jangan pernah mengunci warna UI dengan `D2D1::ColorF(0xFFFFFF)`. Hubungkan komponen dengan konduktor eksternal `ThemeColorEditorText` dan berikan instrumen re-evaluasi `IsDarkMode()` pada *frame render loop* sehingga ia bisa mengubah kulitnya sesuai suhu OS.
  2. **Invisible Boundaries & Emil Polish**: Tombol antarmuka tak selamanya harus dibelenggu border primitif agar eksistensinya valid. Memutilasi paksa garis `stripBorder` dan membuat tombol membaur transparan saat diam (*Rest State*), yang lalu secara reaktif meretas warna latar *Fill Solid* saat diarsir kursor (*Hover/Press*), menciptakan magis ilusi grafis selembut mentega yang ultra-modern.

### Log: 2026-04-12 - Emil Kowalski's Ease-Out in Win32 C++
- **Masalah**: Kurva animasi awal pada komponen desktop bawaan (`EaseInOutCubic`) dengan durasi raksasa (800ms) memberikan efek riak visual yang terasa "*sluggish*" layaknya web era 2010.
- **Penyebab**: Konsep "*Perceived Performance*" web belum diadopsi dalam logika transisi linear C++ Desktop.
- **Pelajaran**: 
  1. *Under 300ms Rule*: Berdasarkan tesis Emil Kowalski, durasi total animasi dipotong masif menjadi 250ms agar mekanisme UI merespons gesit sebelum kesadaran otak manusia mempertanyakan delay-nya.
  2. *Anatomi Matematika Quintic Ease-Out*: Kita memformulasikan ulang ekuasi fisika di `animation_controller.h`. Alih-alih memakai kurva standar `EaseInOut` (yang memulai gerakan secara berlahan dari titik velositas 0), kita memasukkan formula Quintic Out:
     ```cpp
     inline float EaseOutPunchy(float t) {
         float f = (t - 1.0f);
         return f * f * f * f * f + 1.0f; // Ekuivalen ekstrapolasi: 1 - (1 - t)^5
     }
     ```
     **Secara Fisika**: Pada saat pelatuk ditekan (`t = 0`), turunan matematika kurva (velositas/kecepatan) meledak seketika di besaran rotasi absolut 5. Gerakan meloncat seketika layaknya proyektil, langsung mengejar hampir 80% destinasi dalam pecahan milidetik pertama, lalu secara anggun mengurangi kecepatannya hingga statis nol mutlak saat `t = 1`. Inilah rahasia "*Punchy Easing*" (Sensasi Instan Mulus) dalam manipulasi performa UI perseptual yang dicetuskan oleh para *motion engineer* seperti Emil Kowalski & Anime.js.

### Log: 2026-04-12 - Ergonomi "The Soft Renaissance" (Anti-Halation)
- **Masalah**: Penggunaan kontras absolut (#FFF di atas #000) pada editor teks menyebabkan kelelahan mata ekstrem (*Computer Vision Syndrome*).
- **Penyebab**: Fenomena **Halation Effect**. Saat pupil mata melebar di lingkungan gelap, teks putih murni memancarkan cahaya terlalu kuat yang "meluber" ke luar garis huruf, menciptakan pendaran kabur yang melelahkan retina, terutama bagi penderita astigmatisme (silinder).
- **Pelajaran**: 
  1. **Slate & Silver Strategy**: Menurunkan intensitas latar belakang ke **Slate Deep (#121212)** dan teks ke **Soft Silver (#D6D6D6)** secara drastis mengurangi *halation* tanpa mengorbankan keterbacaan.
  2. **Renaissance Balance**: Estetika brutalist tetap bisa dicapai dengan warna "mentah" (slate/ivory) tanpa harus menyakiti mata pengguna. Kenyamanan adalah bagian dari integritas estetika itu sendiri.

## 🦉 Kebijaksanaan (Pelajaran Berharga)

- **Layering Matters**: Selalu tempatkan overlay (seperti lock screen) di urutan terakhir di dalam `Box` root agar mereka memiliki z-index visual tertinggi secara implisit.
- **Surgical Edit**: Saat memperbaiki build failure, lakukan perubahan seminimal mungkin untuk menghindari "comprehension debt" di masa depan.
- **Platform Parity is Math**: Jika kita memahami fisikanya (massa, gaya pegas, gradien cahaya), kita bisa mereplikasi desain web tercanggih sekalipun ke dalam Kotlin native.
- **Invisible Safe Zones**: Selalu sediakan padding statis tak-terlihat di sekitar elemen yang membesar pada kondisi hover, sehingga Layout Shift dapat dihindari tanpa mengorbankan ruang interaksi.

## 💎 Praktik Terbaik

- **Import Precision**: Selalu gunakan path import yang spesifik (misalnya `androidx.compose.ui.draw.zIndex`) daripada import wildcard jika memungkinkan untuk mempercepat build.
- **State Hoisting**: State keamanan (`appUnlocked`, `appLockEnabled`) harus di-hoist ke level teratas agar bisa mengontrol seluruh UI secara konsisten.
- **GraphicsLayer for Performance**: Gunakan `Modifier.graphicsLayer` untuk transformasi 2D/3D karena Compose melakukan optimasi rendering pada level hardware.
- **Asymmetric Shaping**: Gunakan `RoundedCornerShape` asimetris untuk mensimulasikan objek fisik (seperti punggung buku vs ujung halaman) secara visual tanpa memerlukan `Path` yang kompleks.
- **Z-Axis Staggering**: Gunakan translasi Y dengan *offset* yang bertingkat (*staggered*) untuk tumpukan item (seperti notes di dalam folder) guna menciptakan sensasi kedalaman tumpukan fisik di dunia nyata.
- **Matrix Shear (Skew)**: Saat membutuhkan tampilan "miring" yang konsisten tanpa perspektif (distorsi trapesium), gunakan matriks transformasi `[0,1]` (skewX) atau `[1,0]` (skewY). Pastikan untuk mengompensasi pergeseran origin agar objek tetap menempel pada dasarnya.
- **Anti-Halation Principle**: Dalam "Renaissance of Software", teks tidak boleh "berteriak" menyakiti mata. Gunakan warna *off-white* dan latar *off-black* untuk menjaga pupil tetap stabil selama sesi koding panjang.
- **Typography Rhythm (1.2x Line Height)**: Jarak baris yang sedikit lebih lega (1.2x) adalah rahasia untuk mengurangi beban kognitif mata saat membaca struktur logika kode yang padat. Spasi adalah komponen desain, bukan sekadar ruang kosong.
- **Safety Orange Branding**: Warna kontras tinggi (#FF5000) harus digunakan secara fungsional (untuk menandakan fokus/aktif), bukan sekadar hiasan. Ini menciptakan sistem navigasi yang instan dan otoritatif.
