# ctr10.catn.bc
Báo cáo
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hệ thống Tổng hợp Báo cáo Chương trình Hành động số 10 - Điều lệnh CAND</title>

    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: {
                        sans: ['Inter', 'sans-serif'],
                    },
                    colors: {
                        candRed: {
                            50: '#fef2f2',
                            100: '#fee2e2',
                            600: '#dc2626',
                            700: '#b91c1c',
                            800: '#991b1b',
                            900: '#7f1d1d',
                        },
                        candNavy: {
                            800: '#1e293b',
                            900: '#0f172a',
                        },
                        candGold: {
                            400: '#facc15',
                            500: '#eab308',
                        }
                    }
                }
            }
        };
    </script>

    <!-- Google Fonts: Inter -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">

    <!-- SheetJS for Excel Export (.xlsx) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

    <!-- Font Awesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">

    <style>
        body {
            font-family: 'Inter', sans-serif;
            background-color: #f1f5f9;
            color: #0f172a;
        }

        .table-container::-webkit-scrollbar {
            height: 10px;
            width: 10px;
        }

        .table-container::-webkit-scrollbar-track {
            background: #e2e8f0;
            border-radius: 6px;
        }

        .table-container::-webkit-scrollbar-thumb {
            background: #94a3b8;
            border-radius: 6px;
        }

        .table-container::-webkit-scrollbar-thumb:hover {
            background: #64748b;
        }

        .sticky-col-1 {
            position: sticky;
            left: 0;
            z-index: 20;
        }

        .sticky-col-2 {
            position: sticky;
            left: 52px;
            z-index: 20;
        }

        @media print {
            body * {
                visibility: hidden;
            }
            #printArea, #printArea * {
                visibility: visible;
            }
            #printArea {
                position: absolute;
                left: 0;
                top: 0;
                width: 100%;
                background: white;
                padding: 20px;
            }
            .no-print {
                display: none !important;
            }
        }
    </style>
</head>
<body class="min-h-screen flex flex-col">

    <!-- Toast Notification Container -->
    <div id="toastContainer" class="fixed top-4 right-4 z-50 flex flex-col gap-2 max-w-sm w-full pointer-events-none"></div>

    <!-- Header Banner -->
    <header class="bg-gradient-to-r from-candRed-900 via-candRed-800 to-candNavy-900 text-white shadow-lg border-b-4 border-candGold-500">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 sm:py-5">
            <div class="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
                <!-- Title & Emblem -->
                <div class="flex items-start sm:items-center gap-3.5">
                    <div class="w-12 h-12 sm:w-14 sm:h-14 rounded-full bg-candRed-700 border-2 border-candGold-400 flex items-center justify-center shrink-0 shadow-inner">
                        <i class="fa-solid fa-shield-halved text-candGold-400 text-2xl sm:text-3xl"></i>
                    </div>
                    <div>
                        <div class="flex flex-wrap items-center gap-2 text-candGold-400 text-xs sm:text-sm font-semibold uppercase tracking-wider">
                            <span>Công tác Điều lệnh, Quân sự, Võ thuật CAND</span>
                            <span class="hidden sm:inline">•</span>
                            <span id="networkStatusBadge" class="inline-flex items-center gap-1.5 px-2 py-0.5 rounded-full text-[11px] bg-white/10 text-emerald-300 border border-emerald-400/30">
                                <span id="networkDot" class="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
                                <span id="networkStatusText">Đang kiểm tra mạng...</span>
                            </span>
                            <span id="domainBadge" class="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] bg-candNavy-900/80 text-candGold-400 border border-candGold-400/30 font-mono lowercase">
                                <i class="fa-solid fa-globe text-[10px]"></i>
                                <span id="activeDomainHeaderText">localhost</span>
                            </span>
                        </div>
                        <h1 class="text-base sm:text-xl lg:text-2xl font-extrabold tracking-tight leading-snug mt-0.5">
                            HỆ THỐNG TỔNG HỢP BÁO CÁO KẾT QUẢ THỰC HIỆN CHƯƠNG TRÌNH HÀNH ĐỘNG SỐ 10
                        </h1>
                    </div>
                </div>

                <!-- Share Links & Admin Auth Controls -->
                <div class="flex flex-wrap items-center gap-2 self-start lg:self-center">
                    <button type="button" onclick="copyShareLink('unit')" class="inline-flex items-center gap-2 px-3.5 py-2.5 rounded-lg bg-candGold-500 hover:bg-candGold-400 text-candNavy-900 font-bold text-xs sm:text-sm shadow-md transition active:scale-95">
                        <i class="fa-solid fa-link"></i>
                        <span>Link đơn vị (?role=unit)</span>
                    </button>
                    <button type="button" onclick="copyShareLink('admin')" class="inline-flex items-center gap-2 px-3.5 py-2.5 rounded-lg bg-white/15 hover:bg-white/25 text-white border border-white/25 font-bold text-xs sm:text-sm shadow-md transition active:scale-95">
                        <i class="fa-solid fa-user-shield text-candGold-400"></i>
                        <span>Link Admin (?role=admin)</span>
                    </button>

                    <div id="adminAuthControls" class="flex items-center gap-2">
                        <button type="button" id="btnOpenAdminLogin" onclick="openAdminLoginModal()" class="inline-flex items-center gap-2 px-3.5 py-2.5 rounded-lg bg-slate-900 hover:bg-slate-800 text-white border border-candGold-400/50 font-bold text-xs sm:text-sm shadow-md transition active:scale-95">
                            <i class="fa-solid fa-lock text-candGold-400"></i>
                            <span>Đăng nhập Admin</span>
                        </button>

                        <div id="adminLoggedInBadgeGroup" class="hidden flex-wrap items-center gap-2">
                            <span class="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg bg-emerald-500/20 text-emerald-200 border border-emerald-400/40 font-extrabold text-xs sm:text-sm">
                                <i class="fa-solid fa-shield-check text-emerald-300"></i>
                                <span>Chế độ Admin</span>
                            </span>
                            <button type="button" onclick="openChangePinModal()" title="Đổi mã PIN / mật khẩu Admin" class="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg bg-amber-500 hover:bg-amber-400 text-candNavy-900 font-bold text-xs sm:text-sm shadow transition active:scale-95">
                                <i class="fa-solid fa-key"></i>
                                <span>Đổi mã PIN</span>
                            </button>
                            <button type="button" onclick="logoutAdmin()" class="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg bg-rose-700 hover:bg-rose-800 text-white font-bold text-xs sm:text-sm shadow transition active:scale-95">
                                <i class="fa-solid fa-right-from-bracket"></i>
                                <span>Đăng xuất</span>
                            </button>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Navigation Tabs -->
            <div class="mt-5 flex flex-wrap gap-2 border-t border-white/15 pt-4">
                <button type="button" id="tabBtnForm" onclick="switchTab('form')" class="flex items-center gap-2 px-4 py-2.5 rounded-lg font-semibold text-xs sm:text-sm transition bg-white text-candRed-900 shadow-md">
                    <i class="fa-solid fa-file-pen"></i>
                    <span>1. Nhập biểu mẫu báo cáo</span>
                </button>
                <button type="button" id="tabBtnSummary" onclick="switchTab('summary')" class="flex items-center gap-2 px-4 py-2.5 rounded-lg font-semibold text-xs sm:text-sm transition bg-white/10 text-white hover:bg-white/20">
                    <i class="fa-solid fa-table-list"></i>
                    <span>2. Bảng tổng hợp & Xuất Excel</span>
                    <span id="totalReportsBadge" class="ml-1 px-2 py-0.5 text-xs font-bold rounded-full bg-candGold-500 text-candNavy-900">0</span>
                </button>
                <button type="button" id="tabBtnDomain" onclick="switchTab('domain')" class="flex items-center gap-2 px-4 py-2.5 rounded-lg font-semibold text-xs sm:text-sm transition bg-white/10 text-white hover:bg-white/20">
                    <i class="fa-solid fa-globe text-candGold-400"></i>
                    <span>3. Tên miền riêng, Đồng bộ & DNS</span>
                </button>
            </div>
        </div>
    </header>

    <!-- Main Content Area -->
    <main class="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6">

        <!-- ==================== TAB 1: NHẬP BIỂU MẪU BÁO CÁO ==================== -->
        <section id="sectionForm" class="space-y-6">
            <!-- Editing Mode Banner -->
            <div id="editModeBanner" class="hidden bg-amber-50 border-l-4 border-amber-500 p-4 rounded-r-xl shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div class="flex items-center gap-3">
                    <i class="fa-solid fa-pen-to-square text-amber-600 text-xl"></i>
                    <div>
                        <p class="font-bold text-amber-900 text-sm">ĐANG Ở CHẾ ĐỘ CHỈNH SỬA BÁO CÁO</p>
                        <p class="text-xs text-amber-700">Bạn đang cập nhật số liệu cho đơn vị: <strong id="editingUnitLabel"></strong></p>
                    </div>
                </div>
                <button type="button" onclick="cancelEditMode()" class="px-3.5 py-1.5 rounded-lg bg-amber-600 hover:bg-amber-700 text-white text-xs font-semibold transition self-start sm:self-center">
                    Hủy chỉnh sửa / Tạo mới
                </button>
            </div>

            <form id="reportForm" onsubmit="handleFormSubmit(event)" class="space-y-6">
                <input type="hidden" id="editingDocId" value="">

                <!-- Admin Approval Status Field -->
                <div id="adminStatusCard" class="hidden bg-emerald-50 border border-emerald-300 rounded-xl p-4 shadow-2xs">
                    <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                        <div class="flex items-center gap-2.5">
                            <i class="fa-solid fa-stamp text-emerald-700 text-lg"></i>
                            <div>
                                <p class="text-sm font-bold text-emerald-950">Phê duyệt trạng thái báo cáo (Quyền Admin)</p>
                                <p class="text-xs text-emerald-700">Thiết lập trạng thái kiểm duyệt cho báo cáo của đơn vị</p>
                            </div>
                        </div>
                        <select id="reportStatusSelect" class="px-3.5 py-2 rounded-lg border border-emerald-400 bg-white text-slate-900 font-bold text-sm focus:ring-2 focus:ring-emerald-600">
                            <option value="Chờ duyệt">Chờ duyệt</option>
                            <option value="Đã duyệt">Đã duyệt</option>
                            <option value="Cần bổ sung">Cần bổ sung</option>
                        </select>
                    </div>
                </div>

                <!-- Phần I: Thông tin đơn vị gửi -->
                <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                    <div class="bg-candNavy-900 text-white px-5 py-3.5 flex items-center justify-between">
                        <h2 class="font-bold text-sm sm:text-base flex items-center gap-2 uppercase tracking-wide">
                            <i class="fa-solid fa-building-shield text-candGold-400"></i>
                            Phần I. Thông tin đơn vị báo cáo
                        </h2>
                        <span class="text-xs text-slate-300">Các trường có dấu (*) là bắt buộc</span>
                    </div>
                    <div class="p-5 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                        <div class="lg:col-span-2">
                            <label class="block text-sm font-semibold text-slate-700 mb-1.5">
                                Tên đơn vị báo cáo <span class="text-red-600">*</span>
                            </label>
                            <input type="text" id="unitName" required placeholder="VD: Công an phường Phan Đình Phùng / Phòng PX03..."
                                class="w-full px-3.5 py-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 focus:border-candRed-700 text-sm font-medium">
                        </div>
                        <div>
                            <label class="block text-sm font-semibold text-slate-700 mb-1.5">
                                Kỳ báo cáo <span class="text-red-600">*</span>
                            </label>
                            <input type="text" id="reportPeriod" required placeholder="VD: Tháng 11/2026 hoặc Quý IV/2026"
                                class="w-full px-3.5 py-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 focus:border-candRed-700 text-sm">
                        </div>
                        <div>
                            <label class="block text-sm font-semibold text-slate-700 mb-1.5">
                                Cán bộ lập biểu & SĐT
                            </label>
                            <input type="text" id="reporterInfo" placeholder="VD: Đại úy Nguyễn Văn A - 0912..."
                                class="w-full px-3.5 py-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 focus:border-candRed-700 text-sm">
                        </div>
                    </div>
                </div>

                <!-- Phần II: 13 Chỉ tiêu nội dung báo cáo -->
                <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                    <div class="bg-candRed-800 text-white px-5 py-3.5 flex items-center justify-between">
                        <h2 class="font-bold text-sm sm:text-base flex items-center gap-2 uppercase tracking-wide">
                            <i class="fa-solid fa-list-check text-candGold-400"></i>
                            Phần II. Nội dung số liệu thực hiện Chương trình hành động số 10
                        </h2>
                        <span class="text-xs text-red-100 hidden sm:inline">Nhập số "0" nếu không phát sinh</span>
                    </div>

                    <div class="divide-y divide-slate-200">
                        <!-- Mục 1 -> 4 -->
                        <div class="p-5 grid grid-cols-1 md:grid-cols-2 gap-5 bg-slate-50/50">
                            <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
                                <label for="m1_trienKhaiVanBan" class="block text-sm font-bold text-slate-800 mb-2">
                                    1. Số lượt triển khai các văn bản:
                                </label>
                                <div class="flex items-center gap-2">
                                    <input type="number" id="m1_trienKhaiVanBan" min="0" value="0" required
                                        class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 font-semibold text-slate-900">
                                    <span class="text-xs font-medium text-slate-500 shrink-0">lượt</span>
                                </div>
                            </div>

                            <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
                                <label for="m2_thaoLuan" class="block text-sm font-bold text-slate-800 mb-2">
                                    2. Số lượt thảo luận:
                                </label>
                                <div class="flex items-center gap-2">
                                    <input type="number" id="m2_thaoLuan" min="0" value="0" required
                                        class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 font-semibold text-slate-900">
                                    <span class="text-xs font-medium text-slate-500 shrink-0">lượt</span>
                                </div>
                            </div>

                            <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
                                <label for="m3_hoiNghiDoiThoai" class="block text-sm font-bold text-slate-800 mb-2">
                                    3. Số hội nghị đối thoại:
                                </label>
                                <div class="flex items-center gap-2">
                                    <input type="number" id="m3_hoiNghiDoiThoai" min="0" value="0" required
                                        class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 font-semibold text-slate-900">
                                    <span class="text-xs font-medium text-slate-500 shrink-0">hội nghị</span>
                                </div>
                            </div>

                            <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
                                <label for="m4_hoiThaoCuocThi" class="block text-sm font-bold text-slate-800 mb-2">
                                    4. Số hội nghị, hội thảo hoặc cuộc thi:
                                </label>
                                <div class="flex items-center gap-2">
                                    <input type="number" id="m4_hoiThaoCuocThi" min="0" value="0" required
                                        class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 font-semibold text-slate-900">
                                    <span class="text-xs font-medium text-slate-500 shrink-0">cuộc</span>
                                </div>
                            </div>
                        </div>

                        <!-- Mục 5 -->
                        <div class="p-5">
                            <h3 class="text-sm font-bold text-slate-800 mb-3">
                                5. Các sản phẩm công nghệ, phần mềm, ứng dụng có hiệu quả:
                            </h3>
                            <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                                <div>
                                    <label for="m5_soLuongSanPham" class="block text-xs font-semibold text-slate-600 mb-1">Số lượng sản phẩm</label>
                                    <div class="flex items-center gap-2">
                                        <input type="number" id="m5_soLuongSanPham" min="0" value="0" required
                                            class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 font-semibold">
                                        <span class="text-xs text-slate-500 shrink-0">sản phẩm</span>
                                    </div>
                                </div>
                                <div class="md:col-span-2">
                                    <label for="m5_noiDungSanPham" class="block text-xs font-semibold text-slate-600 mb-1">Nội dung chi tiết (nếu có)</label>
                                    <input type="text" id="m5_noiDungSanPham" placeholder="Ghi rõ tên phần mềm, ứng dụng và hiệu quả áp dụng thực tế..."
                                        class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 text-sm">
                                </div>
                            </div>
                        </div>

                        <!-- Mục 6 -->
                        <div class="p-5 bg-slate-50/50">
                            <h3 class="text-sm font-bold text-slate-800 mb-3">
                                6. Số câu lạc bộ đã thành lập:
                            </h3>
                            <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                                <div>
                                    <label for="m6_soLuongCLB" class="block text-xs font-semibold text-slate-600 mb-1">Số lượng câu lạc bộ</label>
                                    <div class="flex items-center gap-2">
                                        <input type="number" id="m6_soLuongCLB" min="0" value="0" required
                                            class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 font-semibold">
                                        <span class="text-xs text-slate-500 shrink-0">CLB</span>
                                    </div>
                                </div>
                                <div class="md:col-span-2">
                                    <label for="m6_tenCLB" class="block text-xs font-semibold text-slate-600 mb-1">Tên câu lạc bộ (nếu có)</label>
                                    <input type="text" id="m6_tenCLB" placeholder="Ghi rõ tên các câu lạc bộ đã thành lập (VD: CLB Võ thuật, CLB Điều lệnh...)"
                                        class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 text-sm">
                                </div>
                            </div>
                        </div>

                        <!-- Mục 7 -->
                        <div class="p-5">
                            <div class="grid grid-cols-1 md:grid-cols-3 gap-4 items-center">
                                <div class="md:col-span-2">
                                    <label for="m7_kiemTraDieuLenh" class="block text-sm font-bold text-slate-800">
                                        7. Số lượt kiểm tra điều lệnh CAND:
                                    </label>
                                    <p class="text-xs text-slate-500 mt-0.5">Tổng số lượt kiểm tra điều lệnh định kỳ và đột xuất của đơn vị</p>
                                </div>
                                <div class="flex items-center gap-2">
                                    <input type="number" id="m7_kiemTraDieuLenh" min="0" value="0" required
                                        class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 font-bold text-slate-900">
                                    <span class="text-xs font-medium text-slate-500 shrink-0">lượt</span>
                                </div>
                            </div>
                        </div>

                        <!-- Mục 8 -->
                        <div class="p-5 bg-red-50/40">
                            <h3 class="text-sm font-bold text-candRed-900 mb-3 flex items-center gap-2">
                                <i class="fa-solid fa-triangle-exclamation text-candRed-700"></i>
                                8. Số lượt xử lý vi phạm điều lệnh CAND:
                            </h3>

                            <div class="space-y-4">
                                <div class="bg-white p-4 rounded-xl border border-red-200">
                                    <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                                        <div>
                                            <label for="m8_viPhamTapThe" class="block text-xs font-bold text-slate-700 mb-1">
                                                - Tập thể (Số lượt xử lý):
                                            </label>
                                            <div class="flex items-center gap-2">
                                                <input type="number" id="m8_viPhamTapThe" min="0" value="0" required
                                                    class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 font-semibold">
                                                <span class="text-xs text-slate-500 shrink-0">lượt</span>
                                            </div>
                                        </div>
                                        <div class="md:col-span-2">
                                            <label for="m8_chiTietTapThe" class="block text-xs font-semibold text-slate-600 mb-1">
                                                Nêu rõ lỗi vi phạm, hình thức xử lý (Tập thể):
                                            </label>
                                            <input type="text" id="m8_chiTietTapThe" placeholder="Ghi rõ tên tập thể, nội dung lỗi vi phạm và hình thức xử lý..."
                                                class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 text-sm">
                                        </div>
                                    </div>
                                </div>

                                <div class="bg-white p-4 rounded-xl border border-red-200 space-y-4">
                                    <div class="flex flex-wrap items-center justify-between gap-2 pb-2 border-b border-slate-100">
                                        <span class="text-xs font-bold text-slate-800 uppercase">
                                            - Cá nhân (Tự động tổng hợp = Lãnh đạo, chỉ huy + CBCS):
                                        </span>
                                        <div class="inline-flex items-center gap-2 px-3 py-1 rounded-lg bg-red-100 text-candRed-900 font-bold text-sm">
                                            <span>Tổng vi phạm cá nhân:</span>
                                            <span id="display_m8_tongCaNhan" class="text-base">0</span>
                                            <span>lượt</span>
                                        </div>
                                    </div>

                                    <div class="grid grid-cols-1 md:grid-cols-3 gap-4 pl-2 border-l-2 border-candRed-600">
                                        <div>
                                            <label for="m8_viPhamLanhDao" class="block text-xs font-bold text-slate-700 mb-1">
                                                + Lãnh đạo, chỉ huy (Số lượt):
                                            </label>
                                            <div class="flex items-center gap-2">
                                                <input type="number" id="m8_viPhamLanhDao" min="0" value="0" required oninput="recalculateTotals()"
                                                    class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 font-semibold">
                                                <span class="text-xs text-slate-500 shrink-0">lượt</span>
                                            </div>
                                        </div>
                                        <div class="md:col-span-2">
                                            <label for="m8_chiTietLanhDao" class="block text-xs font-semibold text-slate-600 mb-1">
                                                Nêu rõ lỗi vi phạm, hình thức xử lý (Lãnh đạo, chỉ huy):
                                            </label>
                                            <input type="text" id="m8_chiTietLanhDao" placeholder="Ghi rõ lỗi vi phạm, hình thức xử lý đối với Lãnh đạo, chỉ huy..."
                                                class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 text-sm">
                                        </div>
                                    </div>

                                    <div class="grid grid-cols-1 md:grid-cols-3 gap-4 pl-2 border-l-2 border-slate-400">
                                        <div>
                                            <label for="m8_viPhamCBCS" class="block text-xs font-bold text-slate-700 mb-1">
                                                + CBCS (Số lượt):
                                            </label>
                                            <div class="flex items-center gap-2">
                                                <input type="number" id="m8_viPhamCBCS" min="0" value="0" required oninput="recalculateTotals()"
                                                    class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 font-semibold">
                                                <span class="text-xs text-slate-500 shrink-0">lượt</span>
                                            </div>
                                        </div>
                                        <div class="md:col-span-2">
                                            <label for="m8_chiTietCBCS" class="block text-xs font-semibold text-slate-600 mb-1">
                                                Nêu rõ lỗi vi phạm, hình thức xử lý (CBCS):
                                            </label>
                                            <input type="text" id="m8_chiTietCBCS" placeholder="Ghi rõ lỗi vi phạm, hình thức xử lý đối với CBCS..."
                                                class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 text-sm">
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <!-- Mục 9 -->
                        <div class="p-5 bg-blue-50/40">
                            <h3 class="text-sm font-bold text-candNavy-900 mb-3 flex items-center gap-2">
                                <i class="fa-solid fa-person-military-pointing text-blue-700"></i>
                                9. Kết quả tập huấn điều lệnh CAND:
                            </h3>

                            <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                                <div class="bg-white p-3.5 rounded-xl border border-slate-200">
                                    <label for="m9_khongThamGia" class="block text-xs font-bold text-slate-700 mb-1.5">
                                        - Số CBCS không tham gia tập huấn theo quy định:
                                    </label>
                                    <div class="flex items-center gap-2">
                                        <input type="number" id="m9_khongThamGia" min="0" value="0" required
                                            class="w-full px-3 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-blue-700 font-semibold">
                                        <span class="text-xs text-slate-500 shrink-0">CBCS</span>
                                    </div>
                                </div>

                                <div class="bg-blue-100/70 p-3.5 rounded-xl border border-blue-300">
                                    <label class="block text-xs font-bold text-blue-950 mb-1.5">
                                        - Số CBCS tham gia tập huấn theo quy định (Tự động):
                                    </label>
                                    <div class="flex items-center justify-between px-3 py-2 rounded-lg bg-white border border-blue-200">
                                        <span id="display_m9_thamGia" class="font-extrabold text-blue-800 text-base">0</span>
                                        <span class="text-xs font-semibold text-blue-600">CBCS (= Đạt + Không đạt)</span>
                                    </div>
                                </div>

                                <div class="bg-white p-3.5 rounded-xl border border-emerald-200">
                                    <label for="m9_datYeuCau" class="block text-xs font-bold text-emerald-800 mb-1.5">
                                        + Số CBCS đạt yêu cầu trở lên:
                                    </label>
                                    <div class="flex items-center gap-2">
                                        <input type="number" id="m9_datYeuCau" min="0" value="0" required oninput="recalculateTotals()"
                                            class="w-full px-3 py-2 rounded-lg border border-emerald-300 focus:ring-2 focus:ring-emerald-600 font-semibold text-emerald-900">
                                        <span class="text-xs text-slate-500 shrink-0">CBCS</span>
                                    </div>
                                </div>

                                <div class="bg-white p-3.5 rounded-xl border border-rose-200">
                                    <label for="m9_khongDatYeuCau" class="block text-xs font-bold text-rose-800 mb-1.5">
                                        + Số CBCS không đạt yêu cầu:
                                    </label>
                                    <div class="flex items-center gap-2">
                                        <input type="number" id="m9_khongDatYeuCau" min="0" value="0" required oninput="recalculateTotals()"
                                            class="w-full px-3 py-2 rounded-lg border border-rose-300 focus:ring-2 focus:ring-rose-600 font-semibold text-rose-900">
                                        <span class="text-xs text-slate-500 shrink-0">CBCS</span>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <!-- Mục 10, 11, 12 -->
                        <div class="p-5 grid grid-cols-1 md:grid-cols-3 gap-5">
                            <div class="bg-slate-50 p-4 rounded-xl border border-slate-200 flex flex-col justify-between">
                                <label for="m10_thongBaoCongKhai" class="block text-sm font-bold text-slate-800 mb-2">
                                    10. Số lượt thông báo công khai kết quả kiểm tra điều lệnh CAND tại các buổi giao ban hoặc Chào cờ định kỳ:
                                </label>
                                <div class="flex items-center gap-2 mt-2">
                                    <input type="number" id="m10_thongBaoCongKhai" min="0" value="0" required
                                        class="w-full px-3.5 py-2 rounded-lg border border-slate-300 bg-white focus:ring-2 focus:ring-candRed-700 font-semibold">
                                    <span class="text-xs font-medium text-slate-500 shrink-0">lượt</span>
                                </div>
                            </div>

                            <div class="bg-slate-50 p-4 rounded-xl border border-slate-200 flex flex-col justify-between">
                                <label for="m11_canBoBanChuyenTrach" class="block text-sm font-bold text-slate-800 mb-2">
                                    11. Số cán bộ bán chuyên trách làm công tác điều lệnh CAND:
                                </label>
                                <div class="flex items-center gap-2 mt-2">
                                    <input type="number" id="m11_canBoBanChuyenTrach" min="0" value="0" required
                                        class="w-full px-3.5 py-2 rounded-lg border border-slate-300 bg-white focus:ring-2 focus:ring-candRed-700 font-semibold">
                                    <span class="text-xs font-medium text-slate-500 shrink-0">đồng chí</span>
                                </div>
                            </div>

                            <div class="bg-slate-50 p-4 rounded-xl border border-slate-200 flex flex-col justify-between">
                                <label for="m12_tuyenTruyenGuong" class="block text-sm font-bold text-slate-800 mb-2">
                                    12. Số lượt tuyên truyền về các gương tập thể, cá nhân tiêu biểu, xuất sắc:
                                </label>
                                <div class="flex items-center gap-2 mt-2">
                                    <input type="number" id="m12_tuyenTruyenGuong" min="0" value="0" required
                                        class="w-full px-3.5 py-2 rounded-lg border border-slate-300 bg-white focus:ring-2 focus:ring-candRed-700 font-semibold">
                                    <span class="text-xs font-medium text-slate-500 shrink-0">lượt</span>
                                </div>
                            </div>
                        </div>

                        <!-- Mục 13 -->
                        <div class="p-5 bg-amber-50/40">
                            <h3 class="text-sm font-bold text-amber-950 mb-3 flex items-center gap-2">
                                <i class="fa-solid fa-award text-amber-600"></i>
                                13. Số tập thể, cá nhân tiêu biểu, xuất sắc được đề nghị khen thưởng trong triển khai thực hiện Chương trình hành động số 10:
                            </h3>
                            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                                <div class="bg-white p-4 rounded-xl border border-amber-200 space-y-3">
                                    <div>
                                        <label for="m13_khenThuongTapThe" class="block text-xs font-bold text-slate-700 mb-1">
                                            - Tập thể (Số lượng):
                                        </label>
                                        <div class="flex items-center gap-2">
                                            <input type="number" id="m13_khenThuongTapThe" min="0" value="0" required
                                                class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-600 font-semibold">
                                            <span class="text-xs text-slate-500 shrink-0">tập thể</span>
                                        </div>
                                    </div>
                                    <div>
                                        <label for="m13_tenTapTheKhenThuong" class="block text-xs font-semibold text-slate-600 mb-1">
                                            Tên tập thể đề nghị khen thưởng (nếu có):
                                        </label>
                                        <input type="text" id="m13_tenTapTheKhenThuong" placeholder="Nhập tên các tập thể được đề nghị khen thưởng..."
                                            class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-600 text-sm">
                                    </div>
                                </div>

                                <div class="bg-white p-4 rounded-xl border border-amber-200 space-y-3">
                                    <div>
                                        <label for="m13_khenThuongCaNhan" class="block text-xs font-bold text-slate-700 mb-1">
                                            - Cá nhân (Số lượng):
                                        </label>
                                        <div class="flex items-center gap-2">
                                            <input type="number" id="m13_khenThuongCaNhan" min="0" value="0" required
                                                class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-600 font-semibold">
                                            <span class="text-xs text-slate-500 shrink-0">cá nhân</span>
                                        </div>
                                    </div>
                                    <div>
                                        <label for="m13_tenCaNhanKhenThuong" class="block text-xs font-semibold text-slate-600 mb-1">
                                            Họ tên cá nhân đề nghị khen thưởng (nếu có):
                                        </label>
                                        <input type="text" id="m13_tenCaNhanKhenThuong" placeholder="Nhập cấp bậc, họ tên các cá nhân đề nghị khen thưởng..."
                                            class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-600 text-sm">
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- Submit Footer -->
                    <div class="bg-slate-100 px-5 py-4 border-t border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-3">
                        <button type="button" onclick="resetFormToDefault()" class="w-full sm:w-auto px-4 py-2.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 font-semibold text-sm transition">
                            <i class="fa-solid fa-rotate-left mr-1.5"></i> Làm mới biểu mẫu
                        </button>
                        <button type="submit" id="btnSubmitReport" class="w-full sm:w-auto px-7 py-3 rounded-lg bg-candRed-800 hover:bg-candRed-900 text-white font-bold text-sm sm:text-base shadow-md transition flex items-center justify-center gap-2 active:scale-95">
                            <i class="fa-solid fa-paper-plane"></i>
                            <span id="btnSubmitText">Gửi báo cáo tổng hợp</span>
                        </button>
                    </div>
                </div>
            </form>
        </section>

        <!-- ==================== TAB 2: BẢNG TỔNG HỢP & XUẤT EXCEL ==================== -->
        <section id="sectionSummary" class="hidden space-y-6">
            <!-- KPI Cards -->
            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
                <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs flex items-center gap-3.5">
                    <div class="w-11 h-11 rounded-lg bg-candNavy-900 text-candGold-400 flex items-center justify-center shrink-0">
                        <i class="fa-solid fa-building-user text-lg"></i>
                    </div>
                    <div>
                        <p class="text-xs font-semibold text-slate-500 uppercase">Đơn vị đã nộp</p>
                        <p id="kpiTotalUnits" class="text-2xl font-extrabold text-slate-900">0</p>
                    </div>
                </div>

                <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs flex items-center gap-3.5">
                    <div class="w-11 h-11 rounded-lg bg-blue-100 text-blue-800 flex items-center justify-center shrink-0">
                        <i class="fa-solid fa-clipboard-check text-lg"></i>
                    </div>
                    <div>
                        <p class="text-xs font-semibold text-slate-500 uppercase">Lượt kiểm tra ĐL</p>
                        <p id="kpiTotalInspections" class="text-2xl font-extrabold text-blue-900">0</p>
                    </div>
                </div>

                <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs flex items-center gap-3.5">
                    <div class="w-11 h-11 rounded-lg bg-red-100 text-candRed-800 flex items-center justify-center shrink-0">
                        <i class="fa-solid fa-gavel text-lg"></i>
                    </div>
                    <div>
                        <p class="text-xs font-semibold text-slate-500 uppercase">Xử lý vi phạm</p>
                        <p id="kpiTotalViolations" class="text-xl font-extrabold text-candRed-800">0 TT / 0 CN</p>
                    </div>
                </div>

                <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs flex items-center gap-3.5">
                    <div class="w-11 h-11 rounded-lg bg-emerald-100 text-emerald-800 flex items-center justify-center shrink-0">
                        <i class="fa-solid fa-user-check text-lg"></i>
                    </div>
                    <div>
                        <p class="text-xs font-semibold text-slate-500 uppercase">Tập huấn đạt YC</p>
                        <p id="kpiTrainingPass" class="text-xl font-extrabold text-emerald-800">0 <span class="text-xs font-semibold text-slate-500">(0%)</span></p>
                    </div>
                </div>

                <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs flex items-center gap-3.5">
                    <div class="w-11 h-11 rounded-lg bg-amber-100 text-amber-700 flex items-center justify-center shrink-0">
                        <i class="fa-solid fa-medal text-lg"></i>
                    </div>
                    <div>
                        <p class="text-xs font-semibold text-slate-500 uppercase">Đề nghị khen thưởng</p>
                        <p id="kpiTotalAwards" class="text-xl font-extrabold text-amber-800">0 TT / 0 CN</p>
                    </div>
                </div>
            </div>

            <!-- Filter & Export Toolbar -->
            <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs flex flex-col lg:flex-row lg:items-center justify-between gap-4">
                <div class="flex flex-col sm:flex-row gap-3 flex-1">
                    <div class="relative flex-1 max-w-md">
                        <i class="fa-solid fa-magnifying-glass absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400 text-sm"></i>
                        <input type="text" id="searchUnitInput" oninput="renderSummaryTable()" placeholder="Tìm kiếm theo tên đơn vị..."
                            class="w-full pl-10 pr-4 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 text-sm">
                    </div>

                    <div class="sm:w-52">
                        <select id="filterPeriodSelect" onchange="renderSummaryTable()"
                            class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 text-sm bg-white font-medium">
                            <option value="ALL">-- Tất cả kỳ báo cáo --</option>
                        </select>
                    </div>

                    <div class="sm:w-48">
                        <select id="filterStatusSelect" onchange="renderSummaryTable()"
                            class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 text-sm bg-white font-medium">
                            <option value="ALL">-- Tất cả trạng thái --</option>
                            <option value="Đã duyệt">Đã duyệt</option>
                            <option value="Chờ duyệt">Chờ duyệt</option>
                            <option value="Cần bổ sung">Cần bổ sung</option>
                        </select>
                    </div>
                </div>

                <div class="flex flex-wrap items-center gap-2.5">
                    <button type="button" id="btnExportExcel" onclick="exportToExcel()" class="inline-flex items-center gap-2 px-4 py-2.5 rounded-lg bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-sm shadow-sm transition active:scale-95">
                        <i class="fa-solid fa-file-excel"></i>
                        <span>Xuất file Excel (.xlsx)</span>
                    </button>
                    <button type="button" id="btnDeleteAllAdmin" onclick="askDeleteAllReports()" class="hidden inline-flex items-center gap-2 px-4 py-2.5 rounded-lg bg-rose-700 hover:bg-rose-800 text-white font-bold text-sm shadow-sm transition active:scale-95">
                        <i class="fa-solid fa-dumpster-fire"></i>
                        <span>Xóa toàn bộ dữ liệu</span>
                    </button>
                </div>
            </div>

            <!-- Non-Admin Notice Banner -->
            <div id="unitRoleNotice" class="bg-slate-100 border border-slate-300 rounded-xl px-4 py-3 text-xs sm:text-sm text-slate-700 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                <div class="flex items-center gap-2">
                    <i class="fa-solid fa-circle-info text-candRed-700"></i>
                    <span>Đang xem ở chế độ <strong>Đơn vị nhập liệu</strong>. Các chức năng phê duyệt trạng thái, chỉnh sửa/xóa bản ghi, xóa toàn bộ hoặc đổi mã PIN yêu cầu <strong>Đăng nhập Admin</strong>.</span>
                </div>
                <button type="button" onclick="openAdminLoginModal()" class="text-candRed-800 hover:text-candRed-950 font-bold underline shrink-0 self-start sm:self-center">
                    Đăng nhập Admin ngay
                </button>
            </div>

            <!-- Summary Table -->
            <div class="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div class="px-5 py-3 bg-candNavy-900 text-white flex items-center justify-between">
                    <h3 class="font-bold text-sm uppercase tracking-wide">
                        Bảng tổng hợp chi tiết 13 chỉ tiêu thực hiện Chương trình hành động số 10
                    </h3>
                    <span class="text-xs text-slate-300 hidden sm:inline">
                        <i class="fa-solid fa-arrows-left-right mr-1"></i> Cuộn ngang để xem đầy đủ 13 nội dung
                    </span>
                </div>

                <div class="table-container overflow-x-auto max-h-[68vh]">
                    <table class="w-full text-xs text-left border-collapse">
                        <thead class="bg-slate-100 text-slate-800 font-bold uppercase sticky top-0 z-30 shadow-2xs">
                            <tr>
                                <th rowspan="2" class="sticky-col-1 bg-slate-200 border border-slate-300 px-3 py-2.5 text-center w-[52px]">STT</th>
                                <th rowspan="2" class="sticky-col-2 bg-slate-200 border border-slate-300 px-3 py-2.5 min-w-[210px]">Đơn vị báo cáo</th>
                                <th rowspan="2" class="border border-slate-300 px-3 py-2.5 min-w-[140px]">Kỳ báo cáo / Ngày gửi</th>
                                <th rowspan="2" class="border border-slate-300 px-3 py-2.5 text-center min-w-[125px]">Trạng thái</th>
                                <th rowspan="2" class="border border-slate-300 px-2.5 py-2.5 text-center min-w-[90px]">1. Triển khai VB</th>
                                <th rowspan="2" class="border border-slate-300 px-2.5 py-2.5 text-center min-w-[85px]">2. Thảo luận</th>
                                <th rowspan="2" class="border border-slate-300 px-2.5 py-2.5 text-center min-w-[90px]">3. HN đối thoại</th>
                                <th rowspan="2" class="border border-slate-300 px-2.5 py-2.5 text-center min-w-[95px]">4. Hội thảo / Cuộc thi</th>
                                <th colspan="2" class="border border-slate-300 px-2.5 py-2 text-center bg-blue-50">5. Sản phẩm CN, PM, ứng dụng</th>
                                <th colspan="2" class="border border-slate-300 px-2.5 py-2 text-center bg-indigo-50">6. Câu lạc bộ đã thành lập</th>
                                <th rowspan="2" class="border border-slate-300 px-2.5 py-2.5 text-center min-w-[95px]">7. Kiểm tra ĐL CAND</th>
                                <th colspan="7" class="border border-slate-300 px-2.5 py-2 text-center bg-red-50 text-candRed-900">8. Xử lý vi phạm điều lệnh CAND</th>
                                <th colspan="4" class="border border-slate-300 px-2.5 py-2 text-center bg-emerald-50 text-emerald-950">9. Kết quả tập huấn ĐL CAND</th>
                                <th rowspan="2" class="border border-slate-300 px-2.5 py-2.5 text-center min-w-[95px]">10. TB công khai</th>
                                <th rowspan="2" class="border border-slate-300 px-2.5 py-2.5 text-center min-w-[95px]">11. CB bán chuyên trách</th>
                                <th rowspan="2" class="border border-slate-300 px-2.5 py-2.5 text-center min-w-[95px]">12. Tuyên truyền gương</th>
                                <th colspan="4" class="border border-slate-300 px-2.5 py-2 text-center bg-amber-50 text-amber-950">13. Đề nghị khen thưởng</th>
                                <th rowspan="2" class="border border-slate-300 px-3 py-2.5 text-center min-w-[155px] bg-slate-200">Thao tác</th>
                            </tr>
                            <tr>
                                <th class="border border-slate-300 px-2 py-2 text-center bg-blue-50/70 min-w-[75px]">SL</th>
                                <th class="border border-slate-300 px-2 py-2 text-left bg-blue-50/70 min-w-[180px]">Nội dung</th>
                                <th class="border border-slate-300 px-2 py-2 text-center bg-indigo-50/70 min-w-[70px]">SL</th>
                                <th class="border border-slate-300 px-2 py-2 text-left bg-indigo-50/70 min-w-[170px]">Tên CLB</th>
                                <th class="border border-slate-300 px-2 py-2 text-center bg-red-50/70 min-w-[75px]">Tập thể</th>
                                <th class="border border-slate-300 px-2 py-2 text-left bg-red-50/70 min-w-[190px]">Lỗi & XL (Tập thể)</th>
                                <th class="border border-slate-300 px-2 py-2 text-center bg-red-100/70 min-w-[80px]">Tổng CN</th>
                                <th class="border border-slate-300 px-2 py-2 text-center bg-red-50/70 min-w-[75px]">LĐ, chỉ huy</th>
                                <th class="border border-slate-300 px-2 py-2 text-left bg-red-50/70 min-w-[190px]">Lỗi & XL (LĐ, chỉ huy)</th>
                                <th class="border border-slate-300 px-2 py-2 text-center bg-red-50/70 min-w-[70px]">CBCS</th>
                                <th class="border border-slate-300 px-2 py-2 text-left bg-red-50/70 min-w-[190px]">Lỗi & XL (CBCS)</th>
                                <th class="border border-slate-300 px-2 py-2 text-center bg-emerald-50/70 min-w-[80px]">Không TG</th>
                                <th class="border border-slate-300 px-2 py-2 text-center bg-emerald-100/70 min-w-[80px]">Tham gia</th>
                                <th class="border border-slate-300 px-2 py-2 text-center bg-emerald-50/70 min-w-[75px]">Đạt YC</th>
                                <th class="border border-slate-300 px-2 py-2 text-center bg-emerald-50/70 min-w-[80px]">Không đạt</th>
                                <th class="border border-slate-300 px-2 py-2 text-center bg-amber-50/70 min-w-[75px]">Tập thể</th>
                                <th class="border border-slate-300 px-2 py-2 text-left bg-amber-50/70 min-w-[170px]">Tên tập thể</th>
                                <th class="border border-slate-300 px-2 py-2 text-center bg-amber-50/70 min-w-[75px]">Cá nhân</th>
                                <th class="border border-slate-300 px-2 py-2 text-left bg-amber-50/70 min-w-[170px]">Họ tên cá nhân</th>
                            </tr>
                        </thead>
                        <tbody id="summaryTableBody" class="divide-y divide-slate-200 bg-white">
                            <!-- Rendered via JavaScript -->
                        </tbody>
                    </table>
                </div>
            </div>
        </section>

        <!-- ==================== TAB 3: TÊN MIỀN RIÊNG, ĐỒNG BỘ & DNS ==================== -->
        <section id="sectionDomain" class="hidden space-y-6">
            <!-- Custom Domain & Share Links Generator -->
            <div class="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div class="bg-candNavy-900 text-white px-5 py-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                    <div>
                        <h2 class="font-bold text-base uppercase tracking-wide flex items-center gap-2">
                            <i class="fa-solid fa-globe text-candGold-400"></i>
                            Cấu hình tên miền riêng (Custom Domain) & Liên kết chia sẻ trực tuyến
                        </h2>
                        <p class="text-xs text-slate-300 mt-0.5">
                            Thiết lập tên miền riêng cho hệ thống (ví dụ: <code class="text-candGold-400">baocao-ct10.example.vn</code>) để tự động sinh liên kết nhập liệu cho các đơn vị và liên kết quản trị cho Admin.
                        </p>
                    </div>
                    <div id="syncModePill" class="px-3 py-1.5 rounded-full text-xs font-bold bg-amber-500/20 text-amber-200 border border-amber-400/40 self-start sm:self-center">
                        Chế độ: Lưu trữ cục bộ (Fallback)
                    </div>
                </div>

                <div class="p-5 space-y-6">
                    <form id="customDomainForm" onsubmit="handleSaveCustomDomain(event)" class="grid grid-cols-1 lg:grid-cols-12 gap-4 items-end bg-slate-50 p-4 rounded-xl border border-slate-200">
                        <div class="lg:col-span-5">
                            <label for="customDomainInput" class="block text-xs font-bold text-slate-700 uppercase mb-1.5">
                                Tên miền riêng hoặc URL triển khai
                            </label>
                            <input type="text" id="customDomainInput" placeholder="VD: baocao-ct10.example.vn hoặc https://baocao-ct10.example.vn"
                                class="w-full px-3.5 py-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 text-sm font-mono bg-white">
                        </div>
                        <div class="lg:col-span-3">
                            <label for="customBasePathInput" class="block text-xs font-bold text-slate-700 uppercase mb-1.5">
                                Thư mục gốc (tùy chọn)
                            </label>
                            <input type="text" id="customBasePathInput" placeholder="VD: / hoặc /ct10/"
                                class="w-full px-3.5 py-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 text-sm font-mono bg-white">
                        </div>
                        <div class="lg:col-span-4 flex flex-wrap gap-2">
                            <button type="submit" class="flex-1 px-4 py-2.5 rounded-lg bg-candRed-800 hover:bg-candRed-900 text-white font-bold text-sm shadow transition flex items-center justify-center gap-2">
                                <i class="fa-solid fa-check"></i>
                                <span>Áp dụng tên miền</span>
                            </button>
                            <button type="button" onclick="useCurrentOriginDomain()" class="px-3.5 py-2.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-100 text-slate-700 font-semibold text-xs transition" title="Dùng tên miền hiện tại của trình duyệt">
                                Dùng domain hiện tại
                            </button>
                        </div>
                    </form>

                    <!-- Generated Unit & Admin Links -->
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div class="p-4 rounded-xl border border-amber-300 bg-amber-50/50 space-y-2.5">
                            <div class="flex items-center justify-between">
                                <span class="text-xs font-extrabold uppercase text-amber-950 flex items-center gap-1.5">
                                    <i class="fa-solid fa-users text-amber-600"></i>
                                    Liên kết cho Đơn vị nhập liệu
                                </span>
                                <span class="text-[11px] font-mono px-2 py-0.5 rounded bg-amber-200/80 text-amber-950 font-bold">?role=unit</span>
                            </div>
                            <div class="flex items-center gap-2">
                                <input type="text" id="generatedUnitLink" readonly
                                    class="w-full px-3 py-2 rounded-lg border border-amber-300 bg-white text-xs font-mono text-slate-800 select-all">
                                <button type="button" onclick="copyShareLink('unit')" class="px-3.5 py-2 rounded-lg bg-candGold-500 hover:bg-candGold-400 text-candNavy-900 font-bold text-xs shrink-0 transition">
                                    <i class="fa-regular fa-copy mr-1"></i> Sao chép
                                </button>
                            </div>
                            <p class="text-[11px] text-amber-900">Gửi link này cho các đơn vị để điền biểu mẫu 13 chỉ tiêu và xem bảng tổng hợp.</p>
                        </div>

                        <div class="p-4 rounded-xl border border-slate-300 bg-slate-50 space-y-2.5">
                            <div class="flex items-center justify-between">
                                <span class="text-xs font-extrabold uppercase text-candNavy-900 flex items-center gap-1.5">
                                    <i class="fa-solid fa-user-shield text-candRed-800"></i>
                                    Liên kết Quản trị viên (Admin)
                                </span>
                                <span class="text-[11px] font-mono px-2 py-0.5 rounded bg-slate-200 text-slate-900 font-bold">?role=admin</span>
                            </div>
                            <div class="flex items-center gap-2">
                                <input type="text" id="generatedAdminLink" readonly
                                    class="w-full px-3 py-2 rounded-lg border border-slate-300 bg-white text-xs font-mono text-slate-800 select-all">
                                <button type="button" onclick="copyShareLink('admin')" class="px-3.5 py-2 rounded-lg bg-candNavy-900 hover:bg-slate-800 text-white font-bold text-xs shrink-0 transition">
                                    <i class="fa-regular fa-copy mr-1"></i> Sao chép
                                </button>
                            </div>
                            <p class="text-[11px] text-slate-600">Mở link này sẽ bật hộp thoại xác thực mã PIN/mật khẩu Admin để duyệt, sửa, xóa và xuất Excel.</p>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Firebase Firestore / Realtime Sync Config Box -->
            <div class="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div class="bg-slate-900 text-white px-5 py-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                    <h3 class="font-bold text-sm sm:text-base uppercase tracking-wide flex items-center gap-2">
                        <i class="fa-solid fa-cloud-arrow-up text-emerald-400"></i>
                        Kết nối Firebase Firestore để đồng bộ dữ liệu đa thiết bị trực tuyến
                    </h3>
                    <span id="firebaseStateLabel" class="text-xs font-mono text-amber-300">Chưa cấu hình khóa riêng (Đang dùng dự phòng)</span>
                </div>

                <div class="p-5 space-y-4">
                    <p class="text-xs sm:text-sm text-slate-600 leading-relaxed">
                        Khi chưa nhập cấu hình Firebase riêng, ứng dụng vẫn hoạt động đầy đủ và lưu dữ liệu trên trình duyệt hiện tại (chế độ dự phòng cục bộ). Để nhiều đơn vị ở các máy tính/điện thoại khác nhau cùng nhập liệu và đồng bộ thời gian thực trên tên miền riêng, hãy dán đối tượng <code class="px-1.5 py-0.5 bg-slate-100 rounded text-slate-900 font-mono text-xs">firebaseConfig</code> từ Firebase Console vào ô dưới đây:
                    </p>

                    <div class="grid grid-cols-1 lg:grid-cols-12 gap-4">
                        <div class="lg:col-span-8">
                            <label for="firebaseConfigJsonInput" class="block text-xs font-bold text-slate-700 uppercase mb-1.5">
                                Firebase Config (JSON hoặc object JS)
                            </label>
                            <textarea id="firebaseConfigJsonInput" rows="5" placeholder='{"apiKey":"...","authDomain":"...","projectId":"...","storageBucket":"...","messagingSenderId":"...","appId":"..."}'
                                class="w-full px-3.5 py-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-emerald-600 text-xs font-mono bg-slate-50"></textarea>
                        </div>
                        <div class="lg:col-span-4 flex flex-col justify-between gap-2.5 bg-slate-50 p-4 rounded-xl border border-slate-200">
                            <div class="space-y-1.5 text-xs text-slate-700">
                                <p class="font-bold text-slate-900">Lưu ý khi bật Cloud Firestore:</p>
                                <p>• Bật <strong>Firestore Database</strong> trong Firebase Console (chế độ Test hoặc Rule cho phép đọc/ghi).</p>
                                <p>• Bật <strong>Anonymous Sign-in</strong> trong Authentication để tự động xác thực phiên nhập liệu.</p>
                                <p>• Thêm tên miền riêng của bạn vào <strong>Authorized domains</strong> trong Firebase Authentication.</p>
                            </div>
                            <div class="flex flex-wrap gap-2 pt-2">
                                <button type="button" onclick="handleSaveFirebaseConfig()" class="flex-1 px-4 py-2 rounded-lg bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs shadow transition">
                                    <i class="fa-solid fa-plug mr-1"></i> Lưu & Kết nối Cloud
                                </button>
                                <button type="button" onclick="handleClearFirebaseConfig()" class="px-3 py-2 rounded-lg border border-slate-300 bg-white hover:bg-rose-50 hover:text-rose-700 text-slate-700 font-semibold text-xs transition">
                                    Xóa cấu hình
                                </button>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- DNS A / CNAME & Hosting Guide -->
            <div class="bg-white rounded-xl border border-slate-200 shadow-sm p-5 space-y-5">
                <h3 class="font-bold text-base text-slate-900 uppercase tracking-wide flex items-center gap-2">
                    <i class="fa-solid fa-server text-candRed-800"></i>
                    Hướng dẫn trỏ DNS (A / CNAME) & Đưa lên Vercel / Netlify / GitHub Pages / Hosting
                </h3>

                <div class="grid grid-cols-1 lg:grid-cols-2 gap-4 text-xs sm:text-sm">
                    <div class="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-2.5">
                        <div class="font-bold text-candNavy-900 flex items-center gap-2">
                            <span class="w-6 h-6 rounded-full bg-candRed-800 text-white text-xs flex items-center justify-center font-extrabold">A</span>
                            <span>Cấu hình bản ghi DNS tại nhà cung cấp tên miền</span>
                        </div>
                        <p class="text-xs text-slate-600 leading-relaxed">
                            Đăng nhập trang quản trị tên miền (tên miền cấp 2 hoặc tên miền phụ như <code class="font-mono">baocao-ct10.example.vn</code>) và thêm bản ghi theo nền tảng hosting:
                        </p>
                        <div class="overflow-x-auto rounded-lg border border-slate-300 bg-white">
                            <table class="w-full text-xs text-left border-collapse">
                                <thead class="bg-slate-100 font-bold text-slate-800">
                                    <tr>
                                        <th class="border-b border-slate-200 px-3 py-2">Nền tảng</th>
                                        <th class="border-b border-slate-200 px-3 py-2">Loại (Type)</th>
                                        <th class="border-b border-slate-200 px-3 py-2">Tên / Host</th>
                                        <th class="border-b border-slate-200 px-3 py-2">Giá trị / Trỏ tới</th>
                                    </tr>
                                </thead>
                                <tbody class="divide-y divide-slate-100 font-mono">
                                    <tr>
                                        <td class="px-3 py-2 font-sans font-semibold">Vercel</td>
                                        <td class="px-3 py-2">A / CNAME</td>
                                        <td class="px-3 py-2">@ hoặc baocao-ct10</td>
                                        <td class="px-3 py-2">76.76.21.21 hoặc cname.vercel-dns.com</td>
                                    </tr>
                                    <tr>
                                        <td class="px-3 py-2 font-sans font-semibold">Netlify</td>
                                        <td class="px-3 py-2">A / CNAME</td>
                                        <td class="px-3 py-2">@ hoặc baocao-ct10</td>
                                        <td class="px-3 py-2">75.2.60.5 hoặc [ten-site].netlify.app</td>
                                    </tr>
                                    <tr>
                                        <td class="px-3 py-2 font-sans font-semibold">GitHub Pages</td>
                                        <td class="px-3 py-2">A / CNAME</td>
                                        <td class="px-3 py-2">@ hoặc baocao-ct10</td>
                                        <td class="px-3 py-2">185.199.108.153 hoặc [user].github.io</td>
                                    </tr>
                                    <tr>
                                        <td class="px-3 py-2 font-sans font-semibold">Hosting riêng</td>
                                        <td class="px-3 py-2">A</td>
                                        <td class="px-3 py-2">@ hoặc baocao-ct10</td>
                                        <td class="px-3 py-2">Địa chỉ IPv4 của máy chủ Hosting</td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>

                    <div class="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-2.5">
                        <div class="font-bold text-candNavy-900 flex items-center gap-2">
                            <span class="w-6 h-6 rounded-full bg-candNavy-900 text-candGold-400 text-xs flex items-center justify-center font-extrabold">B</span>
                            <span>Quy trình xuất bản trực tuyến & kích hoạt SSL</span>
                        </div>
                        <ol class="space-y-2 text-xs text-slate-600 leading-relaxed list-decimal pl-4">
                            <li>Tải tệp <code class="font-mono">index.html</code> này về và tải lên thư mục gốc (<code class="font-mono">public_html</code>, <code class="font-mono">www</code>, repository GitHub hoặc dự án Vercel/Netlify).</li>
                            <li>Trong bảng điều khiển của nền tảng hosting/Vercel/Netlify/GitHub Pages, vào mục <strong>Domains / Custom Domains</strong> và thêm tên miền riêng của bạn.</li>
                            <li>Chờ DNS cập nhật (thường vài phút đến vài giờ) và bật chứng chỉ <strong>HTTPS / SSL</strong> tự động để bảo mật đường truyền.</li>
                            <li>Mở trang trên tên miền mới, vào tab <strong>3. Tên miền riêng, Đồng bộ & DNS</strong>, nhấn <strong>Áp dụng tên miền</strong> rồi sao chép <strong>Link đơn vị (?role=unit)</strong> và <strong>Link Admin (?role=admin)</strong> để sử dụng.</li>
                        </ol>
                    </div>
                </div>
            </div>
        </section>
    </main>

    <!-- ==================== MODAL: ĐĂNG NHẬP ADMIN ==================== -->
    <div id="adminLoginModal" class="fixed inset-0 z-50 hidden bg-slate-900/65 backdrop-blur-xs flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl max-w-md w-full overflow-hidden border border-slate-200">
            <div class="bg-candNavy-900 text-white px-5 py-4 flex items-center justify-between">
                <h3 class="font-bold text-base flex items-center gap-2">
                    <i class="fa-solid fa-user-shield text-candGold-400"></i>
                    <span>Đăng nhập Quản trị viên (Admin)</span>
                </h3>
                <button type="button" onclick="closeAdminLoginModal()" class="text-slate-300 hover:text-white p-1">
                    <i class="fa-solid fa-xmark text-xl"></i>
                </button>
            </div>
            <form onsubmit="handleAdminLogin(event)" class="p-6 space-y-4">
                <div class="bg-amber-50 border border-amber-200 rounded-xl p-3 text-xs text-amber-900 space-y-1">
                    <p class="font-bold"><i class="fa-solid fa-lightbulb text-amber-600 mr-1"></i> Mã PIN / mật khẩu mặc định:</p>
                    <p>Mã PIN: <code class="px-1.5 py-0.5 bg-white rounded border border-amber-300 font-bold">123456</code> hoặc mật khẩu <code class="px-1.5 py-0.5 bg-white rounded border border-amber-300 font-bold">admin2026</code></p>
                </div>

                <div>
                    <label for="adminPinInput" class="block text-sm font-bold text-slate-800 mb-1.5">
                        Nhập mã PIN hoặc mật khẩu Admin
                    </label>
                    <div class="relative">
                        <input type="password" id="adminPinInput" required placeholder="Nhập mã PIN hoặc mật khẩu..."
                            class="w-full px-3.5 py-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 text-sm font-semibold">
                        <button type="button" onclick="togglePasswordVisibility('adminPinInput', 'eyeLoginIcon')" class="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-700">
                            <i id="eyeLoginIcon" class="fa-solid fa-eye"></i>
                        </button>
                    </div>
                    <p id="adminLoginError" class="hidden text-xs font-semibold text-rose-600 mt-1.5">
                        <i class="fa-solid fa-circle-exclamation mr-1"></i> Mã PIN hoặc mật khẩu Admin không chính xác!
                    </p>
                </div>

                <div class="flex justify-end gap-2.5 pt-2">
                    <button type="button" onclick="closeAdminLoginModal()" class="px-4 py-2 rounded-lg border border-slate-300 text-slate-700 font-semibold text-sm hover:bg-slate-50">
                        Hủy
                    </button>
                    <button type="submit" class="px-5 py-2 rounded-lg bg-candRed-800 hover:bg-candRed-900 text-white font-bold text-sm shadow">
                        Xác thực Admin
                    </button>
                </div>
            </form>
        </div>
    </div>

    <!-- ==================== MODAL: ĐỔI MÃ PIN ADMIN ==================== -->
    <div id="changePinModal" class="fixed inset-0 z-50 hidden bg-slate-900/65 backdrop-blur-xs flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl max-w-md w-full overflow-hidden border border-slate-200">
            <div class="bg-candNavy-900 text-white px-5 py-4 flex items-center justify-between">
                <h3 class="font-bold text-base flex items-center gap-2">
                    <i class="fa-solid fa-key text-candGold-400"></i>
                    <span>Đổi mã PIN / mật khẩu Admin</span>
                </h3>
                <button type="button" onclick="closeChangePinModal()" class="text-slate-300 hover:text-white p-1">
                    <i class="fa-solid fa-xmark text-xl"></i>
                </button>
            </div>
            <form onsubmit="handleChangeAdminPin(event)" class="p-6 space-y-4">
                <div>
                    <label for="currentPinInput" class="block text-sm font-bold text-slate-800 mb-1">Mã PIN / mật khẩu hiện tại</label>
                    <input type="password" id="currentPinInput" required placeholder="Nhập mã PIN hiện tại..."
                        class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 text-sm">
                </div>
                <div>
                    <label for="newPinInput" class="block text-sm font-bold text-slate-800 mb-1">Mã PIN / mật khẩu mới (tối thiểu 4 ký tự)</label>
                    <input type="password" id="newPinInput" required minlength="4" placeholder="Nhập mã PIN hoặc mật khẩu mới..."
                        class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 text-sm">
                </div>
                <div>
                    <label for="confirmNewPinInput" class="block text-sm font-bold text-slate-800 mb-1">Xác nhận mã PIN / mật khẩu mới</label>
                    <input type="password" id="confirmNewPinInput" required minlength="4" placeholder="Nhập lại mã PIN mới..."
                        class="w-full px-3.5 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-candRed-700 text-sm">
                </div>
                <p id="changePinError" class="hidden text-xs font-semibold text-rose-600"></p>
                <div class="flex justify-end gap-2.5 pt-2">
                    <button type="button" onclick="closeChangePinModal()" class="px-4 py-2 rounded-lg border border-slate-300 text-slate-700 font-semibold text-sm hover:bg-slate-50">
                        Hủy
                    </button>
                    <button type="submit" class="px-5 py-2 rounded-lg bg-amber-500 hover:bg-amber-400 text-candNavy-900 font-bold text-sm shadow">
                        Lưu mã PIN mới
                    </button>
                </div>
            </form>
        </div>
    </div>

    <!-- ==================== MODAL: XEM CHI TIẾT / IN PHIẾU ==================== -->
    <div id="detailModal" class="fixed inset-0 z-50 hidden bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl max-w-3xl w-full max-h-[90vh] flex flex-col overflow-hidden">
            <div class="bg-candNavy-900 text-white px-5 py-3.5 flex items-center justify-between no-print">
                <h3 id="modalTitle" class="font-bold text-base flex items-center gap-2">
                    <i class="fa-solid fa-file-lines text-candGold-400"></i>
                    <span>Chi tiết phiếu báo cáo đơn vị</span>
                </h3>
                <button type="button" onclick="closeDetailModal()" class="text-slate-300 hover:text-white p-1">
                    <i class="fa-solid fa-xmark text-xl"></i>
                </button>
            </div>

            <div id="printArea" class="p-6 overflow-y-auto space-y-4 text-sm"></div>

            <div class="bg-slate-100 px-5 py-3.5 border-t border-slate-200 flex flex-wrap items-center justify-end gap-2.5 no-print">
                <button type="button" onclick="window.print()" class="px-4 py-2 rounded-lg bg-slate-700 hover:bg-slate-800 text-white font-semibold text-xs sm:text-sm flex items-center gap-2">
                    <i class="fa-solid fa-print"></i> In phiếu báo cáo
                </button>
                <button type="button" onclick="closeDetailModal()" class="px-5 py-2 rounded-lg bg-candRed-800 hover:bg-candRed-900 text-white font-semibold text-xs sm:text-sm">
                    Đóng
                </button>
            </div>
        </div>
    </div>

    <!-- ==================== MODAL: XÁC NHẬN XÓA ==================== -->
    <div id="deleteModal" class="fixed inset-0 z-50 hidden bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl max-w-md w-full p-6 space-y-4">
            <div class="flex items-center gap-3 text-candRed-800">
                <div class="w-10 h-10 rounded-full bg-red-100 flex items-center justify-center shrink-0">
                    <i class="fa-solid fa-trash-can text-lg"></i>
                </div>
                <h3 id="deleteModalTitle" class="font-bold text-base">Xác nhận xóa báo cáo</h3>
            </div>
            <p id="deleteModalMessage" class="text-sm text-slate-600">
                Bạn có chắc chắn muốn xóa báo cáo của đơn vị <strong id="deleteTargetName" class="text-slate-900"></strong> không? Hành động này không thể hoàn tác.
            </p>
            <div class="flex justify-end gap-2.5 pt-2">
                <button type="button" onclick="closeDeleteModal()" class="px-4 py-2 rounded-lg border border-slate-300 text-slate-700 font-semibold text-sm hover:bg-slate-50">
                    Hủy bỏ
                </button>
                <button type="button" id="confirmDeleteBtn" class="px-4 py-2 rounded-lg bg-candRed-700 hover:bg-candRed-800 text-white font-semibold text-sm">
                    Xóa vĩnh viễn
                </button>
            </div>
        </div>
    </div>

    <!-- ==================== APPLICATION LOGIC & CLOUD SYNC ==================== -->
    <script type="module">
        import { initializeApp, getApps, deleteApp } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-app.js";
        import { getAuth, signInAnonymously, signInWithCustomToken, onAuthStateChanged } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-auth.js";
        import { getFirestore, collection, doc, setDoc, addDoc, updateDoc, deleteDoc, onSnapshot, serverTimestamp } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-firestore.js";

        const APP_ID = typeof __app_id !== 'undefined' ? __app_id : 'chuong-trinh-10-cand';
        const COLLECTION_NAME = 'bao_cao_ct10_dieu_lenh';
        const SETTINGS_COLLECTION = 'cau_hinh_he_thong';

        const STORAGE_KEYS = {
            reports: `ct10_reports_fallback_${APP_ID}`,
            pin: `ct10_admin_pin_${APP_ID}`,
            adminSession: `ct10_admin_session_${APP_ID}`,
            domain: `ct10_custom_domain_${APP_ID}`,
            basePath: `ct10_custom_basepath_${APP_ID}`,
            firebaseConfig: `ct10_custom_firebase_config_${APP_ID}`
        };

        const DEFAULT_ADMIN_PINS = ['123456', 'admin2026'];

        let db = null;
        let auth = null;
        let currentUser = null;
        let unsubscribeReports = null;
        let unsubscribePin = null;
        let allReports = [];
        let pendingDeleteId = null;
        let isDeletingAll = false;
        let customAdminPin = null;
        let isAdminLoggedIn = false;

        // ---------------- Toast & Helpers ----------------
        function showToast(message, type = 'success') {
            const container = document.getElementById('toastContainer');
            if (!container) return;
            const toast = document.createElement('div');
            const colors = {
                success: 'bg-emerald-800 text-white border-emerald-500',
                error: 'bg-red-800 text-white border-red-500',
                info: 'bg-slate-800 text-white border-slate-500'
            };
            const icons = {
                success: 'fa-circle-check text-emerald-300',
                error: 'fa-circle-exclamation text-red-300',
                info: 'fa-circle-info text-blue-300'
            };

            toast.className = `pointer-events-auto flex items-center gap-3 px-4 py-3 rounded-xl shadow-lg border-l-4 text-sm font-medium transition-all duration-300 ${colors[type] || colors.info}`;
            toast.innerHTML = `
                <i class="fa-solid ${icons[type] || icons.info} text-base"></i>
                <span class="flex-1">${message}</span>
            `;
            container.appendChild(toast);

            setTimeout(() => {
                toast.style.opacity = '0';
                toast.style.transform = 'translateX(20px)';
                setTimeout(() => toast.remove(), 300);
            }, 3500);
        }

        function escHtml(str) {
            if (str === null || str === undefined) return '';
            return String(str)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;');
        }

        function formatSubmittedTime(r) {
            const sec = r.submittedAtMillis
                || (r.createdAt?.seconds ? r.createdAt.seconds * 1000 : null)
                || (r.updatedAt?.seconds ? r.updatedAt.seconds * 1000 : null);
            if (!sec) return '';
            const d = new Date(sec);
            return `${d.toLocaleDateString('vi-VN')} ${d.toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' })}`;
        }

        function renderStatusBadge(status) {
            const st = status || 'Chờ duyệt';
            if (st === 'Đã duyệt') {
                return `<span class="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300"><i class="fa-solid fa-check"></i> Đã duyệt</span>`;
            }
            if (st === 'Cần bổ sung') {
                return `<span class="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[11px] font-bold bg-amber-100 text-amber-900 border border-amber-300"><i class="fa-solid fa-triangle-exclamation"></i> Cần bổ sung</span>`;
            }
            return `<span class="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[11px] font-bold bg-slate-100 text-slate-700 border border-slate-300"><i class="fa-regular fa-clock"></i> Chờ duyệt</span>`;
        }

        // ---------------- Network Status Monitor ----------------
        function updateNetworkStatusBadge() {
            const isOnline = navigator.onLine !== false;
            const dot = document.getElementById('networkDot');
            const text = document.getElementById('networkStatusText');
            const badge = document.getElementById('networkStatusBadge');
            if (!dot || !text || !badge) return;

            if (isOnline) {
                dot.className = 'w-2 h-2 rounded-full bg-emerald-400 animate-pulse';
                text.textContent = 'Trực tuyến (Online)';
                badge.className = 'inline-flex items-center gap-1.5 px-2 py-0.5 rounded-full text-[11px] bg-white/10 text-emerald-300 border border-emerald-400/30';
            } else {
                dot.className = 'w-2 h-2 rounded-full bg-amber-400';
                text.textContent = 'Ngoại tuyến (Offline)';
                badge.className = 'inline-flex items-center gap-1.5 px-2 py-0.5 rounded-full text-[11px] bg-amber-500/20 text-amber-200 border border-amber-400/30';
            }
        }

        window.addEventListener('online', () => {
            updateNetworkStatusBadge();
            showToast('Kết nối mạng đã khôi phục (Online).', 'success');
        });
        window.addEventListener('offline', () => {
            updateNetworkStatusBadge();
            showToast('Đang mất kết nối mạng, dữ liệu vẫn được lưu cục bộ.', 'info');
        });

        // ---------------- Custom Domain & Share Links ----------------
        function getSavedDomainConfig() {
            let domain = '';
            let basePath = '';
            try {
                domain = localStorage.getItem(STORAGE_KEYS.domain) || '';
                basePath = localStorage.getItem(STORAGE_KEYS.basePath) || '';
            } catch (_) {}

            if (!domain) {
                const loc = window.location;
                domain = loc.host || 'localhost';
                if (!basePath) {
                    const cleanPathname = (loc.pathname || '/').replace(/[^/]*$/, '');
                    basePath = cleanPathname || '/';
                }
            }
            if (!basePath.startsWith('/')) basePath = '/' + basePath;
            if (!basePath.endsWith('/')) basePath = basePath + '/';
            return { domain, basePath };
        }

        function buildShareUrl(roleMode) {
            const { domain, basePath } = getSavedDomainConfig();
            let cleanInput = domain.trim();
            let protocol = window.location.protocol && window.location.protocol.startsWith('http')
                ? window.location.protocol
                : 'https:';

            if (/^https?:\/\//i.test(cleanInput)) {
                try {
                    const parsed = new URL(cleanInput);
                    protocol = parsed.protocol;
                    cleanInput = parsed.host + (parsed.pathname && parsed.pathname !== '/' ? parsed.pathname : '');
                } catch (_) {}
            }

            cleanInput = cleanInput.replace(/\/+$/, '');
            const cleanPath = basePath === '/' ? '/' : basePath;
            const fullBase = cleanInput.includes('/')
                ? `${protocol}//${cleanInput}${cleanInput.endsWith('/') ? '' : '/'}`
                : `${protocol}//${cleanInput}${cleanPath}`;

            const url = new URL(fullBase);
            if (roleMode === 'admin') {
                url.searchParams.set('role', 'admin');
            } else {
                url.searchParams.set('role', 'unit');
            }
            return url.toString();
        }

        function refreshDomainUI() {
            const { domain, basePath } = getSavedDomainConfig();
            const headerDomainEl = document.getElementById('activeDomainHeaderText');
            if (headerDomainEl) {
                headerDomainEl.textContent = domain.replace(/^https?:\/\//i, '').replace(/\/+$/, '');
            }

            const domainInput = document.getElementById('customDomainInput');
            const pathInput = document.getElementById('customBasePathInput');
            if (domainInput && !domainInput.value) domainInput.value = domain;
            if (pathInput && !pathInput.value) pathInput.value = basePath;

            const unitLinkInput = document.getElementById('generatedUnitLink');
            const adminLinkInput = document.getElementById('generatedAdminLink');
            if (unitLinkInput) unitLinkInput.value = buildShareUrl('unit');
            if (adminLinkInput) adminLinkInput.value = buildShareUrl('admin');
        }

        window.handleSaveCustomDomain = function(e) {
            e.preventDefault();
            const rawDomain = (document.getElementById('customDomainInput').value || '').trim();
            let rawPath = (document.getElementById('customBasePathInput').value || '/').trim();
            if (!rawDomain) {
                showToast('Vui lòng nhập tên miền riêng hoặc URL triển khai.', 'error');
                return;
            }
            if (!rawPath.startsWith('/')) rawPath = '/' + rawPath;
            if (!rawPath.endsWith('/')) rawPath = rawPath + '/';

            try {
                localStorage.setItem(STORAGE_KEYS.domain, rawDomain);
                localStorage.setItem(STORAGE_KEYS.basePath, rawPath);
            } catch (_) {}

            refreshDomainUI();
            showToast(`Đã áp dụng tên miền "${rawDomain}" và cập nhật liên kết chia sẻ!`, 'success');
        };

        window.useCurrentOriginDomain = function() {
            try {
                localStorage.removeItem(STORAGE_KEYS.domain);
                localStorage.removeItem(STORAGE_KEYS.basePath);
            } catch (_) {}
            document.getElementById('customDomainInput').value = '';
            document.getElementById('customBasePathInput').value = '';
            refreshDomainUI();
            showToast('Đã khôi phục cấu hình theo tên miền hiện tại.', 'info');
        };

        window.copyShareLink = function(mode = 'unit') {
            const url = buildShareUrl(mode);
            const tempInput = document.createElement('input');
            tempInput.value = url;
            document.body.appendChild(tempInput);
            tempInput.select();
            try {
                document.execCommand('copy');
                if (mode === 'admin') {
                    showToast('Đã sao chép liên kết Admin (?role=admin)!', 'info');
                } else {
                    showToast('Đã sao chép liên kết nhập liệu cho Đơn vị (?role=unit)!', 'success');
                }
            } catch (_) {
                showToast('Không thể sao chép tự động, vui lòng sao chép thủ công.', 'error');
            }
            document.body.removeChild(tempInput);
        };

        // ---------------- Admin Authentication & Role UI ----------------
        function verifyAdminPin(pinValue) {
            const clean = (pinValue || '').trim();
            if (!clean) return false;
            if (customAdminPin) return clean === customAdminPin;
            try {
                const saved = localStorage.getItem(STORAGE_KEYS.pin);
                if (saved) return clean === saved;
            } catch (_) {}
            return DEFAULT_ADMIN_PINS.includes(clean);
        }

        function setSyncedAdminPin(newPin) {
            customAdminPin = String(newPin).trim();
            try {
                localStorage.setItem(STORAGE_KEYS.pin, customAdminPin);
            } catch (_) {}
        }

        function applyRoleUI() {
            const btnLogin = document.getElementById('btnOpenAdminLogin');
            const badgeGroup = document.getElementById('adminLoggedInBadgeGroup');
            const adminStatusCard = document.getElementById('adminStatusCard');
            const btnDeleteAll = document.getElementById('btnDeleteAllAdmin');
            const unitNotice = document.getElementById('unitRoleNotice');

            if (isAdminLoggedIn) {
                btnLogin?.classList.add('hidden');
                badgeGroup?.classList.remove('hidden');
                badgeGroup?.classList.add('flex');
                adminStatusCard?.classList.remove('hidden');
                btnDeleteAll?.classList.remove('hidden');
                unitNotice?.classList.add('hidden');
            } else {
                btnLogin?.classList.remove('hidden');
                badgeGroup?.classList.add('hidden');
                badgeGroup?.classList.remove('flex');
                adminStatusCard?.classList.add('hidden');
                btnDeleteAll?.classList.add('hidden');
                unitNotice?.classList.remove('hidden');
                cancelEditMode();
            }
            renderSummaryTable();
        }

        function restoreAdminSession() {
            try {
                if (sessionStorage.getItem(STORAGE_KEYS.adminSession) === '1') {
                    isAdminLoggedIn = true;
                }
            } catch (_) {}
            applyRoleUI();
        }

        window.openAdminLoginModal = function() {
            document.getElementById('adminPinInput').value = '';
            document.getElementById('adminLoginError').classList.add('hidden');
            document.getElementById('adminLoginModal').classList.remove('hidden');
            setTimeout(() => document.getElementById('adminPinInput')?.focus(), 100);
        };

        window.closeAdminLoginModal = function() {
            document.getElementById('adminLoginModal').classList.add('hidden');
        };

        window.togglePasswordVisibility = function(inputId, iconId) {
            const input = document.getElementById(inputId);
            const icon = document.getElementById(iconId);
            if (!input || !icon) return;
            if (input.type === 'password') {
                input.type = 'text';
                icon.classList.replace('fa-eye', 'fa-eye-slash');
            } else {
                input.type = 'password';
                icon.classList.replace('fa-eye-slash', 'fa-eye');
            }
        };

        window.handleAdminLogin = function(e) {
            e.preventDefault();
            const pin = document.getElementById('adminPinInput').value;
            if (verifyAdminPin(pin)) {
                isAdminLoggedIn = true;
                try {
                    sessionStorage.setItem(STORAGE_KEYS.adminSession, '1');
                } catch (_) {}
                closeAdminLoginModal();
                applyRoleUI();
                showToast('Đã đăng nhập chế độ Quản trị viên (Admin)!', 'success');
            } else {
                document.getElementById('adminLoginError').classList.remove('hidden');
            }
        };

        window.logoutAdmin = function() {
            isAdminLoggedIn = false;
            try {
                sessionStorage.removeItem(STORAGE_KEYS.adminSession);
            } catch (_) {}
            applyRoleUI();
            showToast('Đã đăng xuất Admin, chuyển về chế độ Đơn vị nhập liệu.', 'info');
        };

        window.openChangePinModal = function() {
            if (!isAdminLoggedIn) {
                showToast('Vui lòng đăng nhập Admin để đổi mã PIN!', 'error');
                return;
            }
            document.getElementById('currentPinInput').value = '';
            document.getElementById('newPinInput').value = '';
            document.getElementById('confirmNewPinInput').value = '';
            document.getElementById('changePinError').classList.add('hidden');
            document.getElementById('changePinModal').classList.remove('hidden');
        };

        window.closeChangePinModal = function() {
            document.getElementById('changePinModal').classList.add('hidden');
        };

        window.handleChangeAdminPin = async function(e) {
            e.preventDefault();
            if (!isAdminLoggedIn) return;

            const currentPin = document.getElementById('currentPinInput').value;
            const newPin = document.getElementById('newPinInput').value.trim();
            const confirmPin = document.getElementById('confirmNewPinInput').value.trim();
            const errEl = document.getElementById('changePinError');

            if (!verifyAdminPin(currentPin)) {
                errEl.textContent = 'Mã PIN / mật khẩu hiện tại không đúng.';
                errEl.classList.remove('hidden');
                return;
            }
            if (newPin.length < 4) {
                errEl.textContent = 'Mã PIN / mật khẩu mới phải có ít nhất 4 ký tự.';
                errEl.classList.remove('hidden');
                return;
            }
            if (newPin !== confirmPin) {
                errEl.textContent = 'Xác nhận mã PIN / mật khẩu mới không khớp.';
                errEl.classList.remove('hidden');
                return;
            }

            setSyncedAdminPin(newPin);
            if (db && currentUser) {
                try {
                    const pinDocRef = doc(db, 'artifacts', APP_ID, 'public', 'data', SETTINGS_COLLECTION, 'admin_security');
                    await setDoc(pinDocRef, {
                        adminPin: newPin,
                        updatedAt: serverTimestamp(),
                        updatedBy: currentUser.uid
                    }, { merge: true });
                } catch (err) {
                    console.warn('Sync PIN warning:', err);
                }
            }

            closeChangePinModal();
            showToast('Đã cập nhật mã PIN / mật khẩu Admin thành công!', 'success');
        };

        // ---------------- Tab Switching ----------------
        window.switchTab = function(tabName) {
            const sections = {
                form: document.getElementById('sectionForm'),
                summary: document.getElementById('sectionSummary'),
                domain: document.getElementById('sectionDomain')
            };
            const buttons = {
                form: document.getElementById('tabBtnForm'),
                summary: document.getElementById('tabBtnSummary'),
                domain: document.getElementById('tabBtnDomain')
            };

            Object.keys(sections).forEach(key => {
                if (key === tabName) {
                    sections[key]?.classList.remove('hidden');
                    buttons[key].className = 'flex items-center gap-2 px-4 py-2.5 rounded-lg font-semibold text-xs sm:text-sm transition bg-white text-candRed-900 shadow-md';
                } else {
                    sections[key]?.classList.add('hidden');
                    buttons[key].className = 'flex items-center gap-2 px-4 py-2.5 rounded-lg font-semibold text-xs sm:text-sm transition bg-white/10 text-white hover:bg-white/20';
                }
            });

            if (tabName === 'summary') renderSummaryTable();
            if (tabName === 'domain') refreshDomainUI();
        };

        // ---------------- Auto-Calculation ----------------
        window.recalculateTotals = function() {
            const ld = parseInt(document.getElementById('m8_viPhamLanhDao').value) || 0;
            const cbcs = parseInt(document.getElementById('m8_viPhamCBCS').value) || 0;
            document.getElementById('display_m8_tongCaNhan').textContent = ld + cbcs;

            const dat = parseInt(document.getElementById('m9_datYeuCau').value) || 0;
            const khongDat = parseInt(document.getElementById('m9_khongDatYeuCau').value) || 0;
            document.getElementById('display_m9_thamGia').textContent = dat + khongDat;
        };

        // ---------------- Local Fallback & Firebase Cloud Sync ----------------
        function saveLocalReports(list) {
            try {
                localStorage.setItem(STORAGE_KEYS.reports, JSON.stringify(list));
            } catch (_) {}
        }

        function loadLocalReports() {
            try {
                const raw = localStorage.getItem(STORAGE_KEYS.reports);
                return raw ? JSON.parse(raw) : [];
            } catch (_) {
                return [];
            }
        }

        function parseFirebaseConfigInput(rawText) {
            const trimmed = (rawText || '').trim();
            if (!trimmed) return null;
            try {
                return JSON.parse(trimmed);
            } catch (_) {
                const normalized = trimmed
                    .replace(/^(const|let|var)\s+firebaseConfig\s*=\s*/i, '')
                    .replace(/;\s*$/, '')
                    .replace(/(['"])?([a-zA-Z0-9_]+)(['"])?\s*:/g, '"$2":')
                    .replace(/'/g, '"')
                    .replace(/,\s*}/g, '}');
                return JSON.parse(normalized);
            }
        }

        function loadSavedCustomFirebaseConfig() {
            try {
                const raw = localStorage.getItem(STORAGE_KEYS.firebaseConfig);
                return raw ? JSON.parse(raw) : null;
            } catch (_) {
                return null;
            }
        }

        function updateCloudSyncBadge(isCloudActive, labelText) {
            const pill = document.getElementById('syncModePill');
            const stateLabel = document.getElementById('firebaseStateLabel');
            if (pill) {
                if (isCloudActive) {
                    pill.className = 'px-3 py-1.5 rounded-full text-xs font-bold bg-emerald-500/20 text-emerald-200 border border-emerald-400/40 self-start sm:self-center';
                    pill.textContent = labelText || 'Chế độ: Đồng bộ Cloud trực tuyến';
                } else {
                    pill.className = 'px-3 py-1.5 rounded-full text-xs font-bold bg-amber-500/20 text-amber-200 border border-amber-400/40 self-start sm:self-center';
                    pill.textContent = labelText || 'Chế độ: Lưu trữ cục bộ (Fallback)';
                }
            }
            if (stateLabel) {
                stateLabel.textContent = isCloudActive
                    ? 'Đã kết nối Firebase Firestore (Đồng bộ đa thiết bị)'
                    : 'Chưa cấu hình khóa riêng (Đang dùng dự phòng cục bộ)';
                stateLabel.className = isCloudActive ? 'text-xs font-mono text-emerald-300' : 'text-xs font-mono text-amber-300';
            }
        }

        async function connectFirebaseWithConfig(firebaseConfig) {
            if (typeof unsubscribeReports === 'function') {
                unsubscribeReports();
                unsubscribeReports = null;
            }
            if (typeof unsubscribePin === 'function') {
                unsubscribePin();
                unsubscribePin = null;
            }

            try {
                const existingApp = getApps().find(a => a.name === 'ct10_custom_app');
                if (existingApp) {
                    await deleteApp(existingApp);
                }
            } catch (_) {}

            const app = initializeApp(firebaseConfig, 'ct10_custom_app');
            auth = getAuth(app);
            db = getFirestore(app);

            if (typeof __initial_auth_token !== 'undefined' && __initial_auth_token) {
                await signInWithCustomToken(auth, __initial_auth_token);
            } else {
                await signInAnonymously(auth);
            }

            onAuthStateChanged(auth, (user) => {
                currentUser = user;
                if (user) {
                    updateCloudSyncBadge(true, 'Chế độ: Đồng bộ Cloud trực tuyến');
                    subscribeToReports();
                    subscribeToAdminPin();
                }
            });
        }

        window.handleSaveFirebaseConfig = async function() {
            const raw = document.getElementById('firebaseConfigJsonInput').value;
            let parsed;
            try {
                parsed = parseFirebaseConfigInput(raw);
            } catch (_) {
                showToast('Định dạng firebaseConfig không hợp lệ! Vui lòng kiểm tra lại.', 'error');
                return;
            }
            if (!parsed || !parsed.apiKey || !parsed.projectId) {
                showToast('Thiếu apiKey hoặc projectId trong firebaseConfig.', 'error');
                return;
            }

            try {
                localStorage.setItem(STORAGE_KEYS.firebaseConfig, JSON.stringify(parsed));
            } catch (_) {}

            try {
                await connectFirebaseWithConfig(parsed);
                showToast('Đã kết nối Firebase Firestore và bật đồng bộ trực tuyến!', 'success');
            } catch (err) {
                updateCloudSyncBadge(false, 'Lỗi kết nối Cloud (Dùng cục bộ)');
                showToast('Không thể kết nối Firebase: ' + err.message, 'error');
            }
        };

        window.handleClearFirebaseConfig = async function() {
            try {
                localStorage.removeItem(STORAGE_KEYS.firebaseConfig);
            } catch (_) {}
            document.getElementById('firebaseConfigJsonInput').value = '';
            if (typeof unsubscribeReports === 'function') unsubscribeReports();
            if (typeof unsubscribePin === 'function') unsubscribePin();
            db = null;
            auth = null;
            currentUser = null;
            updateCloudSyncBadge(false, 'Chế độ: Lưu trữ cục bộ (Fallback)');
            showToast('Đã xóa cấu hình Cloud, chuyển về lưu trữ cục bộ.', 'info');
        };

        async function initCloudOrFallback() {
            allReports = loadLocalReports();
            document.getElementById('totalReportsBadge').textContent = allReports.length;
            updatePeriodFilterOptions();
            renderSummaryTable();

            const savedCustom = loadSavedCustomFirebaseConfig();
            const configInput = document.getElementById('firebaseConfigJsonInput');
            if (savedCustom && configInput) {
                configInput.value = JSON.stringify(savedCustom, null, 2);
            }

            try {
                const activeConfig = savedCustom || (typeof __firebase_config !== 'undefined' ? JSON.parse(__firebase_config) : null);
                if (!activeConfig) {
                    throw new Error('No Firebase config provided yet.');
                }
                await connectFirebaseWithConfig(activeConfig);
            } catch (error) {
                console.warn('Running with local storage fallback:', error.message);
                updateCloudSyncBadge(false, 'Chế độ: Lưu trữ cục bộ (Fallback)');
            }
        }

        function subscribeToAdminPin() {
            if (!db || !currentUser) return;
            const pinDocRef = doc(db, 'artifacts', APP_ID, 'public', 'data', SETTINGS_COLLECTION, 'admin_security');
            unsubscribePin = onSnapshot(pinDocRef, (snap) => {
                if (snap.exists() && snap.data().adminPin) {
                    setSyncedAdminPin(snap.data().adminPin);
                }
            }, (err) => console.warn('PIN listener warning:', err));
        }

        function subscribeToReports() {
            if (!db || !currentUser) return;
            const reportsCol = collection(db, 'artifacts', APP_ID, 'public', 'data', COLLECTION_NAME);

            unsubscribeReports = onSnapshot(reportsCol, (snapshot) => {
                const list = [];
                snapshot.forEach((docSnap) => {
                    const data = docSnap.data();
                    const tsMillis = data.submittedAtMillis
                        || (data.updatedAt?.seconds ? data.updatedAt.seconds * 1000 : (data.createdAt?.seconds ? data.createdAt.seconds * 1000 : 0));
                    list.push({ id: docSnap.id, ...data, submittedAtMillis: tsMillis });
                });

                list.sort((a, b) => (b.submittedAtMillis || 0) - (a.submittedAtMillis || 0));
                allReports = list;
                saveLocalReports(list);
                document.getElementById('totalReportsBadge').textContent = allReports.length;
                updatePeriodFilterOptions();
                renderSummaryTable();
            }, (error) => {
                showToast('Lỗi đồng bộ dữ liệu trực tuyến: ' + error.message, 'error');
            });
        }

        function updatePeriodFilterOptions() {
            const select = document.getElementById('filterPeriodSelect');
            if (!select) return;
            const currentVal = select.value;
            const periods = [...new Set(allReports.map(r => (r.reportPeriod || '').trim()).filter(Boolean))];

            select.innerHTML = `<option value="ALL">-- Tất cả kỳ báo cáo (${allReports.length}) --</option>`;
            periods.forEach(p => {
                const count = allReports.filter(r => (r.reportPeriod || '').trim() === p).length;
                const opt = document.createElement('option');
                opt.value = p;
                opt.textContent = `${p} (${count} đơn vị)`;
                select.appendChild(opt);
            });

            if (periods.includes(currentVal)) {
                select.value = currentVal;
            }
        }

        // ---------------- Form Submit & CRUD ----------------
        window.handleFormSubmit = async function(e) {
            e.preventDefault();
            const editingId = document.getElementById('editingDocId').value;
            if (editingId && !isAdminLoggedIn) {
                showToast('Chỉ Quản trị viên (Admin) mới có quyền chỉnh sửa báo cáo đã gửi!', 'error');
                openAdminLoginModal();
                return;
            }

            const btn = document.getElementById('btnSubmitReport');
            const btnText = document.getElementById('btnSubmitText');
            const originalText = btnText.textContent;
            btn.disabled = true;
            btnText.textContent = 'Đang lưu báo cáo...';

            const num = (id) => Math.max(0, parseInt(document.getElementById(id).value) || 0);
            const str = (id) => (document.getElementById(id).value || '').trim();

            const m8_viPhamLanhDao = num('m8_viPhamLanhDao');
            const m8_viPhamCBCS = num('m8_viPhamCBCS');
            const m9_datYeuCau = num('m9_datYeuCau');
            const m9_khongDatYeuCau = num('m9_khongDatYeuCau');

            const existingItem = editingId ? allReports.find(r => r.id === editingId) : null;
            const statusValue = isAdminLoggedIn
                ? (document.getElementById('reportStatusSelect').value || 'Chờ duyệt')
                : (existingItem?.status || 'Chờ duyệt');

            const nowMillis = Date.now();
            const payload = {
                unitName: str('unitName'),
                reportPeriod: str('reportPeriod'),
                reporterInfo: str('reporterInfo'),
                status: statusValue,

                m1_trienKhaiVanBan: num('m1_trienKhaiVanBan'),
                m2_thaoLuan: num('m2_thaoLuan'),
                m3_hoiNghiDoiThoai: num('m3_hoiNghiDoiThoai'),
                m4_hoiThaoCuocThi: num('m4_hoiThaoCuocThi'),

                m5_soLuongSanPham: num('m5_soLuongSanPham'),
                m5_noiDungSanPham: str('m5_noiDungSanPham'),

                m6_soLuongCLB: num('m6_soLuongCLB'),
                m6_tenCLB: str('m6_tenCLB'),

                m7_kiemTraDieuLenh: num('m7_kiemTraDieuLenh'),

                m8_viPhamTapThe: num('m8_viPhamTapThe'),
                m8_chiTietTapThe: str('m8_chiTietTapThe'),
                m8_tongCaNhan: m8_viPhamLanhDao + m8_viPhamCBCS,
                m8_viPhamLanhDao,
                m8_chiTietLanhDao: str('m8_chiTietLanhDao'),
                m8_viPhamCBCS,
                m8_chiTietCBCS: str('m8_chiTietCBCS'),

                m9_khongThamGia: num('m9_khongThamGia'),
                m9_thamGia: m9_datYeuCau + m9_khongDatYeuCau,
                m9_datYeuCau,
                m9_khongDatYeuCau,

                m10_thongBaoCongKhai: num('m10_thongBaoCongKhai'),
                m11_canBoBanChuyenTrach: num('m11_canBoBanChuyenTrach'),
                m12_tuyenTruyenGuong: num('m12_tuyenTruyenGuong'),

                m13_khenThuongTapThe: num('m13_khenThuongTapThe'),
                m13_tenTapTheKhenThuong: str('m13_tenTapTheKhenThuong'),
                m13_khenThuongCaNhan: num('m13_khenThuongCaNhan'),
                m13_tenCaNhanKhenThuong: str('m13_tenCaNhanKhenThuong'),

                submittedAtMillis: existingItem?.submittedAtMillis || nowMillis,
                submittedBy: currentUser?.uid || 'unit-user'
            };

            try {
                if (db && currentUser) {
                    const fsData = { ...payload, updatedAt: serverTimestamp() };
                    if (editingId) {
                        const docRef = doc(db, 'artifacts', APP_ID, 'public', 'data', COLLECTION_NAME, editingId);
                        await updateDoc(docRef, fsData);
                        showToast(`Đã cập nhật báo cáo của "${payload.unitName}"!`, 'success');
                        cancelEditMode();
                    } else {
                        fsData.createdAt = serverTimestamp();
                        const colRef = collection(db, 'artifacts', APP_ID, 'public', 'data', COLLECTION_NAME);
                        await addDoc(colRef, fsData);
                        showToast(`Đã gửi báo cáo của "${payload.unitName}" lên hệ thống!`, 'success');
                        openDetailModal(payload, true);
                        resetFormToDefault();
                    }
                } else {
                    if (editingId) {
                        allReports = allReports.map(r => r.id === editingId ? { ...r, ...payload } : r);
                        saveLocalReports(allReports);
                        showToast(`Đã cập nhật báo cáo của "${payload.unitName}" (lưu cục bộ)!`, 'success');
                        cancelEditMode();
                    } else {
                        const newRec = { id: 'local_' + nowMillis + '_' + Math.random().toString(36).slice(2, 7), ...payload };
                        allReports = [newRec, ...allReports];
                        saveLocalReports(allReports);
                        document.getElementById('totalReportsBadge').textContent = allReports.length;
                        updatePeriodFilterOptions();
                        renderSummaryTable();
                        showToast(`Đã gửi báo cáo của "${payload.unitName}" (lưu cục bộ)!`, 'success');
                        openDetailModal(newRec, true);
                        resetFormToDefault();
                    }
                }
            } catch (error) {
                showToast('Có lỗi xảy ra khi lưu báo cáo: ' + error.message, 'error');
            } finally {
                btn.disabled = false;
                btnText.textContent = originalText;
            }
        };

        window.resetFormToDefault = function() {
            const currentPeriod = document.getElementById('reportPeriod').value;
            document.getElementById('reportForm').reset();
            document.getElementById('editingDocId').value = '';
            document.getElementById('reportPeriod').value = currentPeriod;
            document.getElementById('reportStatusSelect').value = 'Chờ duyệt';
            recalculateTotals();
        };

        window.editReport = function(id) {
            if (!isAdminLoggedIn) {
                showToast('Chỉ Quản trị viên (Admin) mới có quyền chỉnh sửa báo cáo!', 'error');
                openAdminLoginModal();
                return;
            }
            const item = allReports.find(r => r.id === id);
            if (!item) return;

            document.getElementById('editingDocId').value = item.id;
            document.getElementById('unitName').value = item.unitName || '';
            document.getElementById('reportPeriod').value = item.reportPeriod || '';
            document.getElementById('reporterInfo').value = item.reporterInfo || '';
            document.getElementById('reportStatusSelect').value = item.status || 'Chờ duyệt';

            document.getElementById('m1_trienKhaiVanBan').value = item.m1_trienKhaiVanBan || 0;
            document.getElementById('m2_thaoLuan').value = item.m2_thaoLuan || 0;
            document.getElementById('m3_hoiNghiDoiThoai').value = item.m3_hoiNghiDoiThoai || 0;
            document.getElementById('m4_hoiThaoCuocThi').value = item.m4_hoiThaoCuocThi || 0;

            document.getElementById('m5_soLuongSanPham').value = item.m5_soLuongSanPham || 0;
            document.getElementById('m5_noiDungSanPham').value = item.m5_noiDungSanPham || '';

            document.getElementById('m6_soLuongCLB').value = item.m6_soLuongCLB || 0;
            document.getElementById('m6_tenCLB').value = item.m6_tenCLB || '';

            document.getElementById('m7_kiemTraDieuLenh').value = item.m7_kiemTraDieuLenh || 0;

            document.getElementById('m8_viPhamTapThe').value = item.m8_viPhamTapThe || 0;
            document.getElementById('m8_chiTietTapThe').value = item.m8_chiTietTapThe || '';
            document.getElementById('m8_viPhamLanhDao').value = item.m8_viPhamLanhDao || 0;
            document.getElementById('m8_chiTietLanhDao').value = item.m8_chiTietLanhDao || '';
            document.getElementById('m8_viPhamCBCS').value = item.m8_viPhamCBCS || 0;
            document.getElementById('m8_chiTietCBCS').value = item.m8_chiTietCBCS || '';

            document.getElementById('m9_khongThamGia').value = item.m9_khongThamGia || 0;
            document.getElementById('m9_datYeuCau').value = item.m9_datYeuCau || 0;
            document.getElementById('m9_khongDatYeuCau').value = item.m9_khongDatYeuCau || 0;

            document.getElementById('m10_thongBaoCongKhai').value = item.m10_thongBaoCongKhai || 0;
            document.getElementById('m11_canBoBanChuyenTrach').value = item.m11_canBoBanChuyenTrach || 0;
            document.getElementById('m12_tuyenTruyenGuong').value = item.m12_tuyenTruyenGuong || 0;

            document.getElementById('m13_khenThuongTapThe').value = item.m13_khenThuongTapThe || 0;
            document.getElementById('m13_tenTapTheKhenThuong').value = item.m13_tenTapTheKhenThuong || '';
            document.getElementById('m13_khenThuongCaNhan').value = item.m13_khenThuongCaNhan || 0;
            document.getElementById('m13_tenCaNhanKhenThuong').value = item.m13_tenCaNhanKhenThuong || '';

            recalculateTotals();
            document.getElementById('editModeBanner').classList.remove('hidden');
            document.getElementById('editingUnitLabel').textContent = `${item.unitName} (${item.reportPeriod})`;
            document.getElementById('btnSubmitText').textContent = 'Cập nhật báo cáo';
            switchTab('form');
            window.scrollTo({ top: 0, behavior: 'smooth' });
        };

        window.cancelEditMode = function() {
            document.getElementById('editModeBanner').classList.add('hidden');
            document.getElementById('editingDocId').value = '';
            document.getElementById('btnSubmitText').textContent = 'Gửi báo cáo tổng hợp';
            resetFormToDefault();
        };

        window.updateReportStatusQuick = async function(id, newStatus) {
            if (!isAdminLoggedIn) {
                showToast('Chỉ Admin mới có quyền phê duyệt trạng thái báo cáo!', 'error');
                return;
            }
            const item = allReports.find(r => r.id === id);
            if (!item) return;

            try {
                if (db && currentUser && !id.startsWith('local_')) {
                    const docRef = doc(db, 'artifacts', APP_ID, 'public', 'data', COLLECTION_NAME, id);
                    await updateDoc(docRef, { status: newStatus, updatedAt: serverTimestamp() });
                } else {
                    allReports = allReports.map(r => r.id === id ? { ...r, status: newStatus } : r);
                    saveLocalReports(allReports);
                    renderSummaryTable();
                }
                showToast(`Đã cập nhật trạng thái "${item.unitName}" → ${newStatus}`, 'success');
            } catch (err) {
                showToast('Lỗi cập nhật trạng thái: ' + err.message, 'error');
            }
        };

        window.askDeleteReport = function(id) {
            if (!isAdminLoggedIn) {
                showToast('Chỉ Quản trị viên (Admin) mới có quyền xóa báo cáo!', 'error');
                openAdminLoginModal();
                return;
            }
            const item = allReports.find(r => r.id === id);
            if (!item) return;
            isDeletingAll = false;
            pendingDeleteId = id;
            document.getElementById('deleteModalTitle').textContent = 'Xác nhận xóa báo cáo';
            document.getElementById('deleteModalMessage').innerHTML = `Bạn có chắc chắn muốn xóa báo cáo của đơn vị <strong class="text-slate-900">${escHtml(item.unitName)} (${escHtml(item.reportPeriod)})</strong> không? Hành động này không thể hoàn tác.`;
            document.getElementById('deleteModal').classList.remove('hidden');
        };

        window.askDeleteAllReports = function() {
            if (!isAdminLoggedIn) {
                showToast('Chỉ Quản trị viên (Admin) mới có quyền xóa toàn bộ dữ liệu!', 'error');
                openAdminLoginModal();
                return;
            }
            if (allReports.length === 0) {
                showToast('Hiện không có dữ liệu báo cáo nào để xóa.', 'info');
                return;
            }
            isDeletingAll = true;
            pendingDeleteId = null;
            document.getElementById('deleteModalTitle').textContent = 'Cảnh báo: Xóa toàn bộ dữ liệu';
            document.getElementById('deleteModalMessage').innerHTML = `Bạn có chắc chắn muốn xóa <strong class="text-rose-700">TOÀN BỘ ${allReports.length} báo cáo</strong> trên hệ thống không? Hành động này không thể hoàn tác!`;
            document.getElementById('deleteModal').classList.remove('hidden');
        };

        window.closeDeleteModal = function() {
            pendingDeleteId = null;
            isDeletingAll = false;
            document.getElementById('deleteModal').classList.add('hidden');
        };

        document.getElementById('confirmDeleteBtn').addEventListener('click', async () => {
            if (!isAdminLoggedIn) {
                closeDeleteModal();
                return;
            }
            try {
                if (isDeletingAll) {
                    if (db && currentUser) {
                        const targets = allReports.filter(r => !r.id.startsWith('local_'));
                        await Promise.all(targets.map(r => deleteDoc(doc(db, 'artifacts', APP_ID, 'public', 'data', COLLECTION_NAME, r.id))));
                    }
                    allReports = [];
                    saveLocalReports([]);
                    document.getElementById('totalReportsBadge').textContent = '0';
                    updatePeriodFilterOptions();
                    renderSummaryTable();
                    showToast('Đã xóa toàn bộ dữ liệu báo cáo khỏi hệ thống.', 'success');
                } else if (pendingDeleteId) {
                    if (db && currentUser && !pendingDeleteId.startsWith('local_')) {
                        const docRef = doc(db, 'artifacts', APP_ID, 'public', 'data', COLLECTION_NAME, pendingDeleteId);
                        await deleteDoc(docRef);
                    } else {
                        allReports = allReports.filter(r => r.id !== pendingDeleteId);
                        saveLocalReports(allReports);
                        document.getElementById('totalReportsBadge').textContent = allReports.length;
                        updatePeriodFilterOptions();
                        renderSummaryTable();
                    }
                    showToast('Đã xóa báo cáo khỏi hệ thống.', 'success');
                }
            } catch (err) {
                showToast('Lỗi khi xóa dữ liệu: ' + err.message, 'error');
            } finally {
                closeDeleteModal();
            }
        });

        // ---------------- Filtering & Summary Table ----------------
        function getFilteredReports() {
            const search = (document.getElementById('searchUnitInput')?.value || '').toLowerCase().trim();
            const period = document.getElementById('filterPeriodSelect')?.value || 'ALL';
            const status = document.getElementById('filterStatusSelect')?.value || 'ALL';

            return allReports.filter(r => {
                const matchUnit = !search || (r.unitName || '').toLowerCase().includes(search);
                const matchPeriod = period === 'ALL' || (r.reportPeriod || '').trim() === period;
                const matchStatus = status === 'ALL' || (r.status || 'Chờ duyệt') === status;
                return matchUnit && matchPeriod && matchStatus;
            });
        }

        window.renderSummaryTable = function() {
            const filtered = getFilteredReports();
            const tbody = document.getElementById('summaryTableBody');
            if (!tbody) return;

            const sums = {
                m1: 0, m2: 0, m3: 0, m4: 0,
                m5_sl: 0, m6_sl: 0, m7: 0,
                m8_tt: 0, m8_cn: 0, m8_ld: 0, m8_cbcs: 0,
                m9_ktg: 0, m9_tg: 0, m9_dat: 0, m9_kdat: 0,
                m10: 0, m11: 0, m12: 0,
                m13_tt: 0, m13_cn: 0
            };

            filtered.forEach(r => {
                sums.m1 += Number(r.m1_trienKhaiVanBan) || 0;
                sums.m2 += Number(r.m2_thaoLuan) || 0;
                sums.m3 += Number(r.m3_hoiNghiDoiThoai) || 0;
                sums.m4 += Number(r.m4_hoiThaoCuocThi) || 0;
                sums.m5_sl += Number(r.m5_soLuongSanPham) || 0;
                sums.m6_sl += Number(r.m6_soLuongCLB) || 0;
                sums.m7 += Number(r.m7_kiemTraDieuLenh) || 0;

                const ld = Number(r.m8_viPhamLanhDao) || 0;
                const cbcs = Number(r.m8_viPhamCBCS) || 0;
                sums.m8_tt += Number(r.m8_viPhamTapThe) || 0;
                sums.m8_ld += ld;
                sums.m8_cbcs += cbcs;
                sums.m8_cn += (ld + cbcs);

                const dat = Number(r.m9_datYeuCau) || 0;
                const kdat = Number(r.m9_khongDatYeuCau) || 0;
                sums.m9_ktg += Number(r.m9_khongThamGia) || 0;
                sums.m9_dat += dat;
                sums.m9_kdat += kdat;
                sums.m9_tg += (dat + kdat);

                sums.m10 += Number(r.m10_thongBaoCongKhai) || 0;
                sums.m11 += Number(r.m11_canBoBanChuyenTrach) || 0;
                sums.m12 += Number(r.m12_tuyenTruyenGuong) || 0;

                sums.m13_tt += Number(r.m13_khenThuongTapThe) || 0;
                sums.m13_cn += Number(r.m13_khenThuongCaNhan) || 0;
            });

            document.getElementById('kpiTotalUnits').textContent = filtered.length;
            document.getElementById('kpiTotalInspections').textContent = sums.m7.toLocaleString('vi-VN');
            document.getElementById('kpiTotalViolations').textContent = `${sums.m8_tt} TT / ${sums.m8_cn} CN`;
            const passRate = sums.m9_tg > 0 ? Math.round((sums.m9_dat / sums.m9_tg) * 100) : 0;
            document.getElementById('kpiTrainingPass').innerHTML = `${sums.m9_dat.toLocaleString('vi-VN')} <span class="text-xs font-semibold text-slate-500">(${passRate}%)</span>`;
            document.getElementById('kpiTotalAwards').textContent = `${sums.m13_tt} TT / ${sums.m13_cn} CN`;

            if (filtered.length === 0) {
                tbody.innerHTML = `
                    <tr>
                        <td colspan="32" class="px-6 py-12 text-center text-slate-500">
                            <i class="fa-regular fa-folder-open text-3xl text-slate-300 mb-2"></i>
                            <p class="font-medium">Chưa có dữ liệu báo cáo nào phù hợp.</p>
                            <p class="text-xs mt-1">Hãy chuyển sang tab <strong>"Nhập biểu mẫu báo cáo"</strong> để gửi số liệu đầu tiên.</p>
                        </td>
                    </tr>
                `;
                return;
            }

            let html = `
                <tr class="bg-amber-100/90 font-extrabold text-slate-900 border-b-2 border-amber-400">
                    <td class="sticky-col-1 bg-amber-200 border border-slate-300 px-3 py-3 text-center">∑</td>
                    <td class="sticky-col-2 bg-amber-200 border border-slate-300 px-3 py-3 uppercase text-candRed-900">TỔNG CỘNG (${filtered.length} ĐƠN VỊ)</td>
                    <td class="border border-slate-300 px-3 py-3 text-center">—</td>
                    <td class="border border-slate-300 px-3 py-3 text-center">—</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center">${sums.m1}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center">${sums.m2}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center">${sums.m3}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center">${sums.m4}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center bg-blue-100/70">${sums.m5_sl}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center text-slate-400">—</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center bg-indigo-100/70">${sums.m6_sl}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center text-slate-400">—</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center text-blue-900">${sums.m7}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center text-candRed-800">${sums.m8_tt}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center text-slate-400">—</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center bg-red-200/70 text-candRed-900">${sums.m8_cn}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center text-candRed-800">${sums.m8_ld}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center text-slate-400">—</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center text-candRed-800">${sums.m8_cbcs}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center text-slate-400">—</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center">${sums.m9_ktg}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center bg-emerald-200/70 text-emerald-950">${sums.m9_tg}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center text-emerald-800">${sums.m9_dat}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center text-rose-800">${sums.m9_kdat}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center">${sums.m10}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center">${sums.m11}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center">${sums.m12}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center text-amber-900">${sums.m13_tt}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center text-slate-400">—</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center text-amber-900">${sums.m13_cn}</td>
                    <td class="border border-slate-300 px-2.5 py-3 text-center text-slate-400">—</td>
                    <td class="border border-slate-300 px-3 py-3 text-center">—</td>
                </tr>
            `;

            filtered.forEach((r, idx) => {
                const ld = Number(r.m8_viPhamLanhDao) || 0;
                const cbcs = Number(r.m8_viPhamCBCS) || 0;
                const tongCN = ld + cbcs;

                const dat = Number(r.m9_datYeuCau) || 0;
                const kdat = Number(r.m9_khongDatYeuCau) || 0;
                const tongTG = dat + kdat;
                const sentTime = formatSubmittedTime(r);
                const st = r.status || 'Chờ duyệt';

                const statusCellHtml = isAdminLoggedIn
                    ? `<select onchange="updateReportStatusQuick('${r.id}', this.value)" class="px-2 py-1 rounded-lg border text-[11px] font-bold ${
                        st === 'Đã duyệt' ? 'bg-emerald-50 text-emerald-800 border-emerald-300' :
                        st === 'Cần bổ sung' ? 'bg-amber-50 text-amber-900 border-amber-300' :
                        'bg-slate-50 text-slate-700 border-slate-300'
                    }">
                        <option value="Chờ duyệt" ${st === 'Chờ duyệt' ? 'selected' : ''}>Chờ duyệt</option>
                        <option value="Đã duyệt" ${st === 'Đã duyệt' ? 'selected' : ''}>Đã duyệt</option>
                        <option value="Cần bổ sung" ${st === 'Cần bổ sung' ? 'selected' : ''}>Cần bổ sung</option>
                    </select>`
                    : renderStatusBadge(st);

                const actionButtonsHtml = isAdminLoggedIn
                    ? `<div class="inline-flex items-center gap-1.5">
                        <button type="button" onclick="viewReportById('${r.id}')" title="Xem phiếu chi tiết & In" class="p-1.5 rounded bg-slate-100 hover:bg-slate-200 text-slate-700"><i class="fa-solid fa-eye"></i></button>
                        <button type="button" onclick="editReport('${r.id}')" title="Chỉnh sửa số liệu" class="p-1.5 rounded bg-amber-100 hover:bg-amber-200 text-amber-800"><i class="fa-solid fa-pen"></i></button>
                        <button type="button" onclick="askDeleteReport('${r.id}')" title="Xóa báo cáo" class="p-1.5 rounded bg-red-100 hover:bg-red-200 text-red-700"><i class="fa-solid fa-trash"></i></button>
                    </div>`
                    : `<div class="inline-flex items-center gap-1.5">
                        <button type="button" onclick="viewReportById('${r.id}')" title="Xem phiếu chi tiết & In" class="px-2.5 py-1 rounded bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold text-[11px] inline-flex items-center gap-1"><i class="fa-solid fa-eye"></i> Xem</button>
                        <button type="button" onclick="openAdminLoginModal()" title="Đăng nhập Admin để sửa/xóa" class="p-1.5 rounded bg-slate-50 hover:bg-slate-100 text-slate-400"><i class="fa-solid fa-lock"></i></button>
                    </div>`;

                html += `
                    <tr class="hover:bg-slate-50 transition">
                        <td class="sticky-col-1 bg-white border border-slate-300 px-3 py-2.5 text-center font-semibold text-slate-600">${idx + 1}</td>
                        <td class="sticky-col-2 bg-white border border-slate-300 px-3 py-2.5 font-bold text-slate-900">
                            <div>${escHtml(r.unitName)}</div>
                            ${r.reporterInfo ? `<div class="text-[11px] font-normal text-slate-500 mt-0.5">${escHtml(r.reporterInfo)}</div>` : ''}
                        </td>
                        <td class="border border-slate-300 px-3 py-2.5 text-slate-700">
                            <div class="font-semibold">${escHtml(r.reportPeriod)}</div>
                            ${sentTime ? `<div class="text-[10px] text-slate-400 mt-0.5">${escHtml(sentTime)}</div>` : ''}
                        </td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center">${statusCellHtml}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center">${r.m1_trienKhaiVanBan || 0}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center">${r.m2_thaoLuan || 0}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center">${r.m3_hoiNghiDoiThoai || 0}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center">${r.m4_hoiThaoCuocThi || 0}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center font-semibold bg-blue-50/40">${r.m5_soLuongSanPham || 0}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-slate-700">${escHtml(r.m5_noiDungSanPham) || '<span class="text-slate-300">—</span>'}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center font-semibold bg-indigo-50/40">${r.m6_soLuongCLB || 0}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-slate-700">${escHtml(r.m6_tenCLB) || '<span class="text-slate-300">—</span>'}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center font-bold text-blue-900">${r.m7_kiemTraDieuLenh || 0}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center font-semibold text-candRed-800">${r.m8_viPhamTapThe || 0}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-slate-700">${escHtml(r.m8_chiTietTapThe) || '<span class="text-slate-300">—</span>'}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center font-bold bg-red-50/60 text-candRed-900">${tongCN}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center">${ld}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-slate-700">${escHtml(r.m8_chiTietLanhDao) || '<span class="text-slate-300">—</span>'}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center">${cbcs}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-slate-700">${escHtml(r.m8_chiTietCBCS) || '<span class="text-slate-300">—</span>'}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center">${r.m9_khongThamGia || 0}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center font-bold bg-emerald-50/60 text-emerald-950">${tongTG}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center text-emerald-800 font-semibold">${dat}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center text-rose-700 font-semibold">${kdat}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center">${r.m10_thongBaoCongKhai || 0}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center">${r.m11_canBoBanChuyenTrach || 0}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center">${r.m12_tuyenTruyenGuong || 0}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center font-semibold text-amber-900">${r.m13_khenThuongTapThe || 0}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-slate-700">${escHtml(r.m13_tenTapTheKhenThuong) || '<span class="text-slate-300">—</span>'}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center font-semibold text-amber-900">${r.m13_khenThuongCaNhan || 0}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-slate-700">${escHtml(r.m13_tenCaNhanKhenThuong) || '<span class="text-slate-300">—</span>'}</td>
                        <td class="border border-slate-300 px-2.5 py-2.5 text-center">${actionButtonsHtml}</td>
                    </tr>
                `;
            });

            tbody.innerHTML = html;
        };

        // ---------------- Voucher Modal & Print ----------------
        window.viewReportById = function(id) {
            const item = allReports.find(r => r.id === id);
            if (item) openDetailModal(item, false);
        };

        function openDetailModal(r, isJustSubmitted = false) {
            const modal = document.getElementById('detailModal');
            const title = document.getElementById('modalTitle');
            const printArea = document.getElementById('printArea');

            title.innerHTML = isJustSubmitted
                ? `<i class="fa-solid fa-circle-check text-emerald-400"></i> <span>Gửi báo cáo thành công! Phiếu xác nhận số liệu</span>`
                : `<i class="fa-solid fa-file-lines text-candGold-400"></i> <span>Chi tiết báo cáo: ${escHtml(r.unitName)}</span>`;

            const tongCN = (Number(r.m8_viPhamLanhDao) || 0) + (Number(r.m8_viPhamCBCS) || 0);
            const tongTG = (Number(r.m9_datYeuCau) || 0) + (Number(r.m9_khongDatYeuCau) || 0);

            printArea.innerHTML = `
                <div class="text-center border-b-2 border-slate-800 pb-4">
                    <p class="font-bold uppercase text-xs tracking-wider text-slate-600">BÁO CÁO KẾT QUẢ THỰC HIỆN CHƯƠNG TRÌNH HÀNH ĐỘNG SỐ 10</p>
                    <h4 class="text-lg font-extrabold uppercase text-slate-900 mt-1">ĐƠN VỊ: ${escHtml(r.unitName)}</h4>
                    <p class="text-xs font-semibold text-slate-600 mt-1">Kỳ báo cáo: ${escHtml(r.reportPeriod)} ${r.reporterInfo ? ` | Người lập biểu: ${escHtml(r.reporterInfo)}` : ''}</p>
                </div>

                <table class="w-full border-collapse border border-slate-400 text-xs sm:text-sm">
                    <thead>
                        <tr class="bg-slate-100 font-bold">
                            <th class="border border-slate-400 px-3 py-2 w-12 text-center">TT</th>
                            <th class="border border-slate-400 px-3 py-2 text-left">Nội dung báo cáo</th>
                            <th class="border border-slate-400 px-3 py-2 text-left w-1/2">Số liệu & Diễn giải chi tiết</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr><td class="border border-slate-400 px-3 py-2 text-center font-bold">1</td><td class="border border-slate-400 px-3 py-2">Số lượt triển khai các văn bản</td><td class="border border-slate-400 px-3 py-2 font-bold">${r.m1_trienKhaiVanBan || 0} lượt</td></tr>
                        <tr><td class="border border-slate-400 px-3 py-2 text-center font-bold">2</td><td class="border border-slate-400 px-3 py-2">Số lượt thảo luận</td><td class="border border-slate-400 px-3 py-2 font-bold">${r.m2_thaoLuan || 0} lượt</td></tr>
                        <tr><td class="border border-slate-400 px-3 py-2 text-center font-bold">3</td><td class="border border-slate-400 px-3 py-2">Số hội nghị đối thoại</td><td class="border border-slate-400 px-3 py-2 font-bold">${r.m3_hoiNghiDoiThoai || 0} hội nghị</td></tr>
                        <tr><td class="border border-slate-400 px-3 py-2 text-center font-bold">4</td><td class="border border-slate-400 px-3 py-2">Số hội nghị, hội thảo hoặc cuộc thi</td><td class="border border-slate-400 px-3 py-2 font-bold">${r.m4_hoiThaoCuocThi || 0} cuộc</td></tr>
                        <tr>
                            <td class="border border-slate-400 px-3 py-2 text-center font-bold">5</td>
                            <td class="border border-slate-400 px-3 py-2">Các sản phẩm công nghệ, phần mềm, ứng dụng có hiệu quả</td>
                            <td class="border border-slate-400 px-3 py-2">
                                <strong>${r.m5_soLuongSanPham || 0}</strong> sản phẩm
                                ${r.m5_noiDungSanPham ? `<div class="text-slate-700 mt-1">- Nội dung: ${escHtml(r.m5_noiDungSanPham)}</div>` : ''}
                            </td>
                        </tr>
                        <tr>
                            <td class="border border-slate-400 px-3 py-2 text-center font-bold">6</td>
                            <td class="border border-slate-400 px-3 py-2">Số câu lạc bộ đã thành lập</td>
                            <td class="border border-slate-400 px-3 py-2">
                                <strong>${r.m6_soLuongCLB || 0}</strong> CLB
                                ${r.m6_tenCLB ? `<div class="text-slate-700 mt-1">- Tên CLB: ${escHtml(r.m6_tenCLB)}</div>` : ''}
                            </td>
                        </tr>
                        <tr><td class="border border-slate-400 px-3 py-2 text-center font-bold">7</td><td class="border border-slate-400 px-3 py-2">Số lượt kiểm tra điều lệnh CAND</td><td class="border border-slate-400 px-3 py-2 font-bold">${r.m7_kiemTraDieuLenh || 0} lượt</td></tr>
                        <tr>
                            <td class="border border-slate-400 px-3 py-2 text-center font-bold">8</td>
                            <td class="border border-slate-400 px-3 py-2">Số lượt xử lý vi phạm điều lệnh CAND</td>
                            <td class="border border-slate-400 px-3 py-2 space-y-1">
                                <div>- Tập thể: <strong>${r.m8_viPhamTapThe || 0}</strong> lượt ${r.m8_chiTietTapThe ? `(${escHtml(r.m8_chiTietTapThe)})` : ''}</div>
                                <div>- Cá nhân: <strong>${tongCN}</strong> lượt, trong đó:</div>
                                <div class="pl-3">+ Lãnh đạo, chỉ huy: <strong>${r.m8_viPhamLanhDao || 0}</strong> lượt ${r.m8_chiTietLanhDao ? `(${escHtml(r.m8_chiTietLanhDao)})` : ''}</div>
                                <div class="pl-3">+ CBCS: <strong>${r.m8_viPhamCBCS || 0}</strong> lượt ${r.m8_chiTietCBCS ? `(${escHtml(r.m8_chiTietCBCS)})` : ''}</div>
                            </td>
                        </tr>
                        <tr>
                            <td class="border border-slate-400 px-3 py-2 text-center font-bold">9</td>
                            <td class="border border-slate-400 px-3 py-2">Kết quả tập huấn điều lệnh CAND</td>
                            <td class="border border-slate-400 px-3 py-2 space-y-1">
                                <div>- Số CBCS không tham gia tập huấn theo quy định: <strong>${r.m9_khongThamGia || 0}</strong> CBCS</div>
                                <div>- Số CBCS tham gia tập huấn theo quy định: <strong>${tongTG}</strong> CBCS, trong đó:</div>
                                <div class="pl-3">+ Số CBCS đạt yêu cầu trở lên: <strong>${r.m9_datYeuCau || 0}</strong> CBCS</div>
                                <div class="pl-3">+ Số CBCS không đạt yêu cầu: <strong>${r.m9_khongDatYeuCau || 0}</strong> CBCS</div>
                            </td>
                        </tr>
                        <tr><td class="border border-slate-400 px-3 py-2 text-center font-bold">10</td><td class="border border-slate-400 px-3 py-2">Số lượt thông báo công khai kết quả kiểm tra điều lệnh CAND tại giao ban / Chào cờ định kỳ</td><td class="border border-slate-400 px-3 py-2 font-bold">${r.m10_thongBaoCongKhai || 0} lượt</td></tr>
                        <tr><td class="border border-slate-400 px-3 py-2 text-center font-bold">11</td><td class="border border-slate-400 px-3 py-2">Số cán bộ bán chuyên trách làm công tác điều lệnh CAND</td><td class="border border-slate-400 px-3 py-2 font-bold">${r.m11_canBoBanChuyenTrach || 0} đồng chí</td></tr>
                        <tr><td class="border border-slate-400 px-3 py-2 text-center font-bold">12</td><td class="border border-slate-400 px-3 py-2">Số lượt tuyên truyền về các gương tập thể, cá nhân tiêu biểu, xuất sắc</td><td class="border border-slate-400 px-3 py-2 font-bold">${r.m12_tuyenTruyenGuong || 0} lượt</td></tr>
                        <tr>
                            <td class="border border-slate-400 px-3 py-2 text-center font-bold">13</td>
                            <td class="border border-slate-400 px-3 py-2">Số tập thể, cá nhân tiêu biểu, xuất sắc được đề nghị khen thưởng</td>
                            <td class="border border-slate-400 px-3 py-2 space-y-1">
                                <div>- Tập thể: <strong>${r.m13_khenThuongTapThe || 0}</strong> ${r.m13_tenTapTheKhenThuong ? `(${escHtml(r.m13_tenTapTheKhenThuong)})` : ''}</div>
                                <div>- Cá nhân: <strong>${r.m13_khenThuongCaNhan || 0}</strong> ${r.m13_tenCaNhanKhenThuong ? `(${escHtml(r.m13_tenCaNhanKhenThuong)})` : ''}</div>
                            </td>
                        </tr>
                    </tbody>
                </table>
            `;

            modal.classList.remove('hidden');
        }

        window.closeDetailModal = function() {
            document.getElementById('detailModal').classList.add('hidden');
        };

        // ---------------- Export to Excel (.xlsx) ----------------
        window.exportToExcel = function() {
            const filtered = getFilteredReports();
            if (filtered.length === 0) {
                showToast('Không có dữ liệu để xuất file Excel!', 'error');
                return;
            }

            const headers = [
                'STT', 'Đơn vị báo cáo', 'Kỳ báo cáo', 'Ngày giờ gửi', 'Trạng thái', 'Cán bộ lập biểu',
                '1. Số lượt triển khai các văn bản', '2. Số lượt thảo luận', '3. Số hội nghị đối thoại',
                '4. Số hội nghị, hội thảo hoặc cuộc thi', '5. Sản phẩm CN, PM, ứng dụng (Số lượng)',
                '5. Sản phẩm CN, PM, ứng dụng (Nội dung)', '6. Câu lạc bộ đã thành lập (Số lượng)',
                '6. Tên câu lạc bộ', '7. Số lượt kiểm tra điều lệnh CAND', '8. Vi phạm Tập thể (Số lượt)',
                '8. Vi phạm Tập thể (Lỗi vi phạm, hình thức xử lý)', '8. Vi phạm Cá nhân (Tổng số lượt)',
                '8. Vi phạm Lãnh đạo, chỉ huy (Số lượt)', '8. Vi phạm Lãnh đạo, chỉ huy (Lỗi vi phạm, hình thức xử lý)',
                '8. Vi phạm CBCS (Số lượt)', '8. Vi phạm CBCS (Lỗi vi phạm, hình thức xử lý)',
                '9. Tập huấn: Số CBCS không tham gia theo quy định', '9. Tập huấn: Số CBCS tham gia theo quy định',
                '9. Tập huấn: Số CBCS đạt yêu cầu trở lên', '9. Tập huấn: Số CBCS không đạt yêu cầu',
                '10. Số lượt thông báo công khai kết quả kiểm tra ĐL', '11. Số cán bộ bán chuyên trách công tác ĐL CAND',
                '12. Số lượt tuyên truyền gương tiêu biểu, xuất sắc', '13. Đề nghị khen thưởng Tập thể (Số lượng)',
                '13. Tên tập thể đề nghị khen thưởng', '13. Đề nghị khen thưởng Cá nhân (Số lượng)',
                '13. Họ tên cá nhân đề nghị khen thưởng'
            ];

            const sums = Array(headers.length).fill('');
            sums[0] = 'TỔNG CỘNG';
            sums[1] = `${filtered.length} đơn vị`;

            const sumCols = [6, 7, 8, 9, 10, 12, 14, 15, 17, 18, 20, 22, 23, 24, 25, 26, 27, 28, 29, 31];
            sumCols.forEach(c => sums[c] = 0);

            const rows = filtered.map((r, index) => {
                const ld = Number(r.m8_viPhamLanhDao) || 0;
                const cbcs = Number(r.m8_viPhamCBCS) || 0;
                const tongCN = ld + cbcs;

                const dat = Number(r.m9_datYeuCau) || 0;
                const kdat = Number(r.m9_khongDatYeuCau) || 0;
                const tongTG = dat + kdat;

                const row = [
                    index + 1,
                    r.unitName || '',
                    r.reportPeriod || '',
                    formatSubmittedTime(r),
                    r.status || 'Chờ duyệt',
                    r.reporterInfo || '',
                    Number(r.m1_trienKhaiVanBan) || 0,
                    Number(r.m2_thaoLuan) || 0,
                    Number(r.m3_hoiNghiDoiThoai) || 0,
                    Number(r.m4_hoiThaoCuocThi) || 0,
                    Number(r.m5_soLuongSanPham) || 0,
                    r.m5_noiDungSanPham || '',
                    Number(r.m6_soLuongCLB) || 0,
                    r.m6_tenCLB || '',
                    Number(r.m7_kiemTraDieuLenh) || 0,
                    Number(r.m8_viPhamTapThe) || 0,
                    r.m8_chiTietTapThe || '',
                    tongCN,
                    ld,
                    r.m8_chiTietLanhDao || '',
                    cbcs,
                    r.m8_chiTietCBCS || '',
                    Number(r.m9_khongThamGia) || 0,
                    tongTG,
                    dat,
                    kdat,
                    Number(r.m10_thongBaoCongKhai) || 0,
                    Number(r.m11_canBoBanChuyenTrach) || 0,
                    Number(r.m12_tuyenTruyenGuong) || 0,
                    Number(r.m13_khenThuongTapThe) || 0,
                    r.m13_tenTapTheKhenThuong || '',
                    Number(r.m13_khenThuongCaNhan) || 0,
                    r.m13_tenCaNhanKhenThuong || ''
                ];

                sumCols.forEach(c => {
                    sums[c] += Number(row[c]) || 0;
                });

                return row;
            });

            const sheetData = [
                ['BẢNG TỔNG HỢP KẾT QUẢ THỰC HIỆN CHƯƠNG TRÌNH HÀNH ĐỘNG SỐ 10 (CÔNG TÁC ĐIỀU LỆNH CAND)'],
                [`Ngày xuất báo cáo: ${new Date().toLocaleDateString('vi-VN')}`],
                [],
                headers,
                sums,
                ...rows
            ];

            const ws = XLSX.utils.aoa_to_sheet(sheetData);
            ws['!cols'] = headers.map((h, i) => {
                if (i === 0) return { wch: 8 };
                if (i === 1) return { wch: 28 };
                if ([2, 3, 4, 5].includes(i)) return { wch: 18 };
                if ([11, 13, 16, 19, 21, 30, 32].includes(i)) return { wch: 30 };
                return { wch: 16 };
            });

            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Tong_Hop_CT10');
            XLSX.writeFile(wb, `Tong_Hop_Bao_Cao_CT10_DieuLenh_${new Date().toISOString().slice(0, 10)}.xlsx`);
            showToast('Đã xuất file Excel (.xlsx) thành công!', 'success');
        };

        // ---------------- Initialization ----------------
        window.addEventListener('DOMContentLoaded', () => {
            const now = new Date();
            const month = String(now.getMonth() + 1).padStart(2, '0');
            const year = now.getFullYear();
            const periodInput = document.getElementById('reportPeriod');
            if (periodInput && !periodInput.value) {
                periodInput.value = `Tháng ${month}/${year}`;
            }

            updateNetworkStatusBadge();
            refreshDomainUI();
            recalculateTotals();
            restoreAdminSession();
            initCloudOrFallback();

            const params = new URLSearchParams(window.location.search);
            if (params.get('role') === 'admin' && !isAdminLoggedIn) {
                openAdminLoginModal();
            }
        });
    </script>
</body>
</html>
