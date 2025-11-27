let currentLang = 'vi';
const sidebar = document.getElementById('sidebar');
const body = document.body;

// --- Dữ liệu ngôn ngữ (Đã cập nhật) ---
const langData = {
    'vi': {
        title: "Hướng dẫn sử dụng K-UTILS PRO v2.3",
        zaloLinkText: "Gửi phản hồi qua Zalo",
        nav: [
            { href: "#intro", icon: "fas fa-home", text: "Tổng quan", group: false },
            { href: "#install", icon: "fas fa-cog", text: "Cài đặt & Sửa lỗi", group: false },
            { group: true, text: "Tính năng chính" },
            { href: "#finance", icon: "fas fa-coins", text: "Tài chính & Kế toán", group: false },
            { href: "#email", icon: "fas fa-envelope", text: "Email Marketing", group: false },
            { href: "#data", icon: "fas fa-database", text: "Xử lý Dữ liệu", group: false },
            { href: "#advanced", icon: "fas fa-magic", text: "Nâng cao", group: false },
            { href: "#utils", icon: "fas fa-tools", text: "Tiện ích khác", group: false },
            { group: true, text: "Hỗ trợ" },
            { href: "#faq", icon: "fas fa-question-circle", text: "FAQ", group: false },
            { href: "#donate", icon: "fas fa-coffee", text: "Mời Cafe", group: false, accent: true },
        ],
        intro_content: `
            <h1>K-Utils PRO <span style="font-size: 0.5em; background: #e0f2fe; color: #0284c7; padding: 4px 8px; border-radius: 4px; vertical-align: middle;">v2.3.0</span></h1>
            <p>Giải pháp Excel Add-in toàn diện cho dân văn phòng. Tự động hóa công việc nhàm chán chỉ với 1 cú click.</p>
            <div style="background: #eff6ff; border: 1px solid #3b82f6; padding: 25px; border-radius: 12px; text-align: center; margin-top: 20px;">
                <a href="https://drive.google.com/drive/folders/1xHwO7JacMJNvxg8jLfyOdrzhVvDKQKA8" target="_blank" style="display: inline-block; background: #2563eb; color: white; padding: 12px 30px; border-radius: 30px; text-decoration: none; font-weight: bold; box-shadow: 0 4px 10px rgba(37,99,235,0.3);">
                    <i class="fas fa-download"></i> Tải Bộ Cài Đặt (.exe)
                </a>
                <div style="margin-top: 15px; font-size: 13px; color: #475569;">
                    <i class="fab fa-windows"></i> Hỗ trợ Windows 10/11 | Excel 2010 - 2024 (32/64-bit)
                </div>
            </div>
        `,
        install_content: `
            <h2><i class="fas fa-cog"></i> Hướng dẫn cài đặt (Quan trọng)</h2>
            <div class="alert-box">
                <strong><i class="fas fa-exclamation-triangle"></i> Lưu ý:</strong> Vui lòng đóng tất cả file Excel trước khi cài đặt.
            </div>
            
            <h3>Bước 1: Chạy file cài đặt</h3>
            <p>Tải file <code>Setup_KUtils_Pro.exe</code> và chạy. Chọn ngôn ngữ <strong>Tiếng Việt</strong> và bấm Next liên tục cho đến khi hoàn tất.</p>

            <h3>Bước 2: Sửa lỗi không hiện Tab (Rất hay gặp)</h3>
            <p>Do chính sách bảo mật của Microsoft, đôi khi Add-in bị chặn. Nếu bạn mở Excel mà không thấy tab K-Utils PRO, hãy làm theo 2 cách sau:</p>
            
            <h4>Cách 1: Bật thủ công trong Excel</h4>
            <ol>
                <li>Vào <strong>File</strong> > <strong>Options</strong> > <strong>Add-ins</strong>.</li>
                <li>Ở mục <em>Manage</em> chọn <strong>Excel Add-ins</strong> > Bấm <strong>Go...</strong></li>
                <li>Tích chọn vào ô <strong>K-Utils PRO</strong> > OK.</li>
            </ol>

            <h4>Cách 2: Unblock file (Nếu Cách 1 không được)</h4>
            <ol>
                <li>Truy cập thư mục cài đặt: <code>%AppData%\\Microsoft\\AddIns\\</code></li>
                <li>Chuột phải vào file <code>KUtils_Pro.xlam</code> > Chọn <strong>Properties</strong>.</li>
                <li>Ở tab General, tích vào ô <strong>Unblock</strong> (ở dưới cùng) > Apply > OK.</li>
                <li>Khởi động lại Excel.</li>
            </ol>
        `,
        finance_content: `
            <h2><i class="fas fa-coins"></i> 1. Tài chính & Kế toán</h2>
            <p>Tính năng đọc số tiền thành chữ hỗ trợ đa vùng miền và ngoại tệ.</p>

            <h3>Hàm =VND (Đọc tiền Việt)</h3>
            <table class="excel-table">
                <tr><th>Cú pháp</th><th>Ví dụ</th><th>Kết quả hiển thị</th></tr>
                <tr>
                    <td><code>=VND(A1)</code></td>
                    <td>125,000</td>
                    <td>Một trăm hai mươi lăm <strong>nghìn</strong> đồng chẵn. (Giọng Bắc)</td>
                </tr>
                <tr>
                    <td><code>=VND(A1, 1)</code></td>
                    <td>125,000</td>
                    <td>Một trăm hai mươi lăm <strong>ngàn</strong> đồng chẵn. (Giọng Nam)</td>
                </tr>
                <tr>
                    <td><code>=VND(A1, 0, 2)</code></td>
                    <td>125,000</td>
                    <td>MỘT TRĂM HAI MƯƠI LĂM NGHÌN ĐỒNG CHẴN. (Viết hoa hết)</td>
                </tr>
            </table>

            <h3>Hàm =DOC_TIEN (Đọc ngoại tệ)</h3>
            <p>Cú pháp: <code>=DOC_TIEN(Số, "Tên tiền", "Tên lẻ")</code></p>
            <ul>
                <li>Ví dụ: <code>=DOC_TIEN(A1, "Euro", "xen")</code> -> <em>Một trăm Euro và mười xen.</em></li>
                <li>Ví dụ: <code>=DOC_TIEN(A1, "Yên", "")</code> -> <em>Một nghìn Yên chẵn.</em></li>
            </ul>
        `,
        email_content: `
            <h2><i class="fas fa-envelope"></i> 2. Email Marketing & Gửi Lương</h2>
            <p>Gửi Email hàng loạt qua Outlook với nội dung và file đính kèm riêng biệt cho từng người. Cực kỳ hữu ích để gửi <strong>Phiếu lương</strong>.</p>

            <div class="alert-info">
                <strong><i class="fas fa-lightbulb"></i> Nguyên tắc hoạt động:</strong> Bạn chuẩn bị dữ liệu trong Excel, sau đó bôi đen Cột Email và bấm nút. Phần mềm sẽ tự động lấy thông tin ở các cột bên cạnh để điền vào Email.
            </div>

            <h3>Cấu trúc bảng dữ liệu (Bắt buộc)</h3>
            <p>Hãy sắp xếp dữ liệu đúng theo thứ tự 5 cột như hình dưới đây:</p>

            <table class="excel-table">
                <tr>
                    <th class="col-header">A</th>
                    <th class="col-header cell-selected" style="border-bottom: 2px solid #2563eb;">B (Chọn cột này)</th>
                    <th class="col-header">C</th>
                    <th class="col-header">D</th>
                    <th class="col-header">E</th>
                </tr>
                <tr>
                    <th>Tên Người Nhận</th>
                    <th style="color:#2563eb;">Địa Chỉ Email</th>
                    <th>Tiêu Đề Thư</th>
                    <th>Nội Dung Thư</th>
                    <th>Đường Dẫn File</th>
                </tr>
                <tr>
                    <td>Nguyễn Văn A</td>
                    <td class="cell-selected">anv@gmail.com</td>
                    <td>Phiếu lương T10</td>
                    <td>Chào A, lương tháng này...</td>
                    <td>D:\\Luong\\Phieu_A.pdf</td>
                </tr>
                <tr>
                    <td>Trần Thị B</td>
                    <td class="cell-selected">btt@yahoo.com</td>
                    <td>Thông báo nợ</td>
                    <td>Chào B, bạn còn nợ...</td>
                    <td>(Để trống nếu ko có)</td>
                </tr>
            </table>

            <h3>Cách thực hiện:</h3>
            <ol>
                <li>Chuẩn bị bảng dữ liệu như trên.</li>
                <li><strong>QUAN TRỌNG:</strong> Bôi đen vùng chứa địa chỉ Email (Ví dụ <code>B2:B3</code>). Không bôi đen cả bảng, chỉ bôi cột Email.</li>
                <li>Bấm nút <strong>Gửi Email Lương</strong> trên Ribbon K-Utils.</li>
                <li>Outlook sẽ bật lên và tạo các email nháp (Draft). Bạn kiểm tra lại rồi bấm Gửi.</li>
            </ol>
        `,
        data_content: `
            <h2><i class="fas fa-database"></i> 3. Xử lý Dữ liệu</h2>
            
            <h3>Gộp File Excel (Merge)</h3>
            <p>Gom dữ liệu từ nhiều file Excel con vào 1 file tổng duy nhất.</p>
            <ul>
                <li>Bấm nút <strong>Gộp File</strong> > Chọn thư mục chứa file.</li>
                <li>Phần mềm sẽ tự động mở từng file và copy dữ liệu vào file hiện tại.</li>
            </ul>

            <h3>Hàm RegEx (Tách dữ liệu khó)</h3>
            <p>Dùng để lấy số điện thoại hoặc Email từ chuỗi văn bản lộn xộn.</p>
            <ul>
                <li>Lấy số: <code>=RegExExtract(A1, "\\d+")</code></li>
                <li>Lấy Email: <code>=RegExExtract(A1, "[a-zA-Z0-9._-]+@[a-z]+\\.[a-z]+")</code></li>
            </ul>
        `,
        advanced_content: `
            <h2><i class="fas fa-magic"></i> 4. Tính năng Nâng cao (v2.3)</h2>
            
            <h3>Tách File (Split Workbook)</h3>
            <p>Chia nhỏ 1 sheet lớn thành nhiều file con dựa trên cột (Ví dụ: Tách lương theo Phòng ban).</p>
            <ol>
                <li>Bấm <strong>Tách File</strong>.</li>
                <li>Bước 1: Quét chọn toàn bộ bảng dữ liệu.</li>
                <li>Bước 2: Chọn 1 ô trong cột muốn tách (VD: Cột Phòng ban).</li>
                <li>Kết quả: Các file con sẽ được tạo trong thư mục mới.</li>
            </ol>

            <h3>Chuyển mã Font (TCVN3 -> Unicode)</h3>
            <div class="alert-danger">
                <strong><i class="fas fa-exclamation-circle"></i> Cảnh báo:</strong> Tính năng này đôi khi không hoàn hảo 100%. Vui lòng **SAO LƯU FILE GỐC** trước khi thực hiện chuyển mã!
            </div>
            <p>Sửa lỗi font chữ bị loằng ngoằng (dạng .VnTime) khi mở file cũ.</p>
            <ul>
                <li>Bôi đen vùng lỗi font.</li>
                <li>Bấm <strong>Chuyển Mã Font</strong> > Chọn nguồn TCVN3 (ABC).</li>
            </ul>
        `,
        utils_content: `
            <h2><i class="fas fa-tools"></i> 5. Tiện ích khác</h2>
            
            <div class="feature-grid">
                <div class="feature-card">
                    <i class="fas fa-qrcode feature-icon"></i>
                    <strong>Tạo QR Code</strong>
                    <p style="font-size:13px; margin-top:5px;">Chọn ô chứa link/text > Bấm nút Tạo QR. Mã QR sẽ được chèn ngay bên cạnh.</p>
                </div>
                <div class="feature-card">
                    <i class="fas fa-list feature-icon"></i>
                    <strong>Tạo Mục Lục</strong>
                    <p style="font-size:13px; margin-top:5px;">Tự động tạo sheet "Muc Luc" chứa đường dẫn (Link) đến tất cả các sheet khác trong file.</p>
                </div>
                <div class="feature-card">
                    <i class="fas fa-eye feature-icon"></i>
                    <strong>Hiện Sheet Ẩn</strong>
                    <p style="font-size:13px; margin-top:5px;">Excel mặc định chỉ cho Unhide từng sheet. Nút này giúp hiện tất cả cùng lúc.</p>
                </div>
                <div class="feature-card">
                    <i class="fas fa-calendar-alt feature-icon"></i>
                    <strong>Sửa Lỗi Ngày</strong>
                    <p style="font-size:13px; margin-top:5px;">Chuyển ngày dạng text (26.11.2023) thành dạng Date chuẩn để tính toán.</p>
                </div>
            </div>
        `,
        faq_content: `
            <h2><i class="fas fa-question-circle"></i> Câu hỏi thường gặp</h2>
            <ul>
                <li><strong>Q: Tại sao tôi cài xong không thấy Tab?</strong><br>A: Xem lại mục 1 (Cài đặt) ở trên, làm theo Cách 2 (Unblock file).</li>
                <li><strong>Q: Phần mềm có virus không?</strong><br>A: Cam kết 100% sạch. Đây là mã nguồn mở VBA.</li>
                <li><strong>Q: Có dùng được trên Macbook không?</strong><br>A: Không, chỉ hỗ trợ Windows.</li>
            </ul>
        `,
        donate_content: `
            <div style="text-align: center;">
                <h2 style="border:none; justify-content: center;"><i class="fas fa-mug-hot"></i> Mời tác giả một ly cà phê?</h2>
                <p>Phần mềm hoàn toàn miễn phí. Sự ủng hộ của bạn là động lực để mình phát triển tiếp!</p>
                
                <div class="bank-card">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <span style="font-weight:800; font-size:18px;">TECHCOMBANK</span>
                        <i class="fas fa-wifi" style="transform: rotate(90deg);"></i>
                    </div>
                    <div style="margin-top:20px;">Account Number</div>
                    <div class="bank-num" id="accNum">1907 5093 4460 12</div>
                    <div style="display:flex; justify-content:space-between; align-items:flex-end;">
                        <div>
                            <div style="font-size:12px; opacity:0.8;">CARD HOLDER</div>
                            <div style="font-weight:600; font-size:16px;">TRAN TUAN KHOA</div>
                        </div>
                        <button class="copy-btn" onclick="copyToClipboard()"><i class="far fa-copy"></i></button>
                    </div>
                </div>
            </div>
        `,
        footer_content: `&copy; 2025 K-Utils PRO. Developed by Tran Tuan Khoa.`,
        toast_copied: "Đã sao chép số tài khoản!",
    },
    'en': {
        title: "K-UTILS PRO v2.3 User Manual",
        zaloLinkText: "Send Feedback via Zalo",
        nav: [
            { href: "#intro", icon: "fas fa-home", text: "Overview", group: false },
            { href: "#install", icon: "fas fa-cog", text: "Installation & Fixes", group: false },
            { group: true, text: "Main Features" },
            { href: "#finance", icon: "fas fa-coins", text: "Finance & Accounting", group: false },
            { href: "#email", icon: "fas fa-envelope", text: "Email Marketing", group: false },
            { href: "#data", icon: "fas fa-database", text: "Data Processing", group: false },
            { href: "#advanced", icon: "fas fa-magic", text: "Advanced", group: false },
            { href: "#utils", icon: "fas fa-tools", text: "Other Utilities", group: false },
            { group: true, text: "Support" },
            { href: "#faq", icon: "fas fa-question-circle", text: "FAQ", group: false },
            { href: "#donate", icon: "fas fa-coffee", text: "Buy Author a Coffee", group: false, accent: true },
        ],
        intro_content: `
            <h1>K-Utils PRO <span style="font-size: 0.5em; background: #e0f2fe; color: #0284c7; padding: 4px 8px; border-radius: 4px; vertical-align: middle;">v2.3.0</span></h1>
            <p>The comprehensive Excel Add-in solution for office workers. Automate tedious tasks with just one click.</p>
            <div style="background: #eff6ff; border: 1px solid #3b82f6; padding: 25px; border-radius: 12px; text-align: center; margin-top: 20px;">
                <a href="https://drive.google.com/drive/folders/1xHwO7JacMJNvxg8jLfyOdrzhVvDKQKA8" target="_blank" style="display: inline-block; background: #2563eb; color: white; padding: 12px 30px; border-radius: 30px; text-decoration: none; font-weight: bold; box-shadow: 0 4px 10px rgba(37,99,235,0.3);">
                    <i class="fas fa-download"></i> Download Installer (.exe)
                </a>
                <div style="margin-top: 15px; font-size: 13px; color: #475569;">
                    <i class="fab fa-windows"></i> Supports Windows 10/11 | Excel 2010 - 2024 (32/64-bit)
                </div>
            </div>
        `,
        install_content: `
            <h2><i class="fas fa-cog"></i> Installation Guide (Important)</h2>
            <div class="alert-box">
                <strong><i class="fas fa-exclamation-triangle"></i> Note:</strong> Please close all Excel files before installation.
            </div>
            
            <h3>Step 1: Run the installation file</h3>
            <p>Download and run the <code>Setup_KUtils_Pro.exe</code> file. Select the language (Vietnamese) and click Next until completion.</p>

            <h3>Step 2: Fix the "Tab Not Showing" error (Very Common)</h3>
            <p>Due to Microsoft's security policy, the Add-in is sometimes blocked. If you open Excel and don't see the K-Utils PRO tab, follow these 2 methods:</p>
            
            <h4>Method 1: Manual activation in Excel</h4>
            <ol>
                <li>Go to <strong>File</strong> > <strong>Options</strong> > <strong>Add-ins</strong>.</li>
                <li>In the <em>Manage</em> section, select <strong>Excel Add-ins</strong> > Click <strong>Go...</strong></li>
                <li>Check the box for <strong>K-Utils PRO</strong> > OK.</li>
            </ol>

            <h4>Method 2: Unblock the file (If Method 1 fails)</h4>
            <ol>
                <li>Access the installation folder: <code>%AppData%\\Microsoft\\AddIns\\</code></li>
                <li>Right-click the file <code>KUtils_Pro.xlam</code> > Select <strong>Properties</strong>.</li>
                <li>In the General tab, check the <strong>Unblock</strong> box (at the bottom) > Apply > OK.</li>
                <li>Restart Excel.</li>
            </ol>
        `,
        finance_content: `
            <h2><i class="fas fa-coins"></i> 1. Finance & Accounting</h2>
            <p>The feature to read amounts into words supports multiple regions and foreign currencies.</p>

            <h3>=VND Function (Read Vietnamese Currency)</h3>
            <table class="excel-table">
                <tr><th>Syntax</th><th>Example</th><th>Result Display</th></tr>
                <tr>
                    <td><code>=VND(A1)</code></td>
                    <td>125,000</td>
                    <td>One hundred twenty-five <strong>thousand</strong> dong only. (Northern accent)</td>
                </tr>
                <tr>
                    <td><code>=VND(A1, 1)</code></td>
                    <td>125,000</td>
                    <td>One hundred twenty-five <strong>thousand</strong> dong only. (Southern accent)</td>
                </tr>
                <tr>
                    <td><code>=VND(A1, 0, 2)</code></td>
                    <td>125,000</td>
                    <td>ONE HUNDRED TWENTY-FIVE THOUSAND DONG ONLY. (All caps)</td>
                </tr>
            </table>

            <h3>=DOC_TIEN Function (Read Foreign Currency)</h3>
            <p>Syntax: <code>=DOC_TIEN(Amount, "Currency Name", "Minor Unit Name")</code></p>
            <ul>
                <li>Example: <code>=DOC_TIEN(A1, "Euro", "cent")</code> -> <em>One hundred Euro and ten cents.</em></li>
                <li>Example: <code>=DOC_TIEN(A1, "Yen", "")</code> -> <em>One thousand Yen only.</em></li>
            </ul>
        `,
        email_content: `
            <h2><i class="fas fa-envelope"></i> 2. Email Marketing & Payslip Sending</h2>
            <p>Send bulk emails via Outlook with separate content and attachments for each recipient. Extremely useful for sending <strong>Payslips</strong>.</p>

            <div class="alert-info">
                <strong><i class="fas fa-lightbulb"></i> Principle of Operation:</strong> You prepare the data in Excel, then select the Email Column and click the button. The software automatically retrieves information from adjacent columns to populate the email.
            </div>

            <h3>Data Table Structure (Required)</h3>
            <p>Please arrange the data in the exact order of 5 columns as shown below:</p>

            <table class="excel-table">
                <tr>
                    <th class="col-header">A</th>
                    <th class="col-header cell-selected" style="border-bottom: 2px solid #2563eb;">B (Select this column)</th>
                    <th class="col-header">C</th>
                    <th class="col-header">D</th>
                    <th class="col-header">E</th>
                </tr>
                <tr>
                    <th>Recipient Name</th>
                    <th style="color:#2563eb;">Email Address</th>
                    <th>Subject</th>
                    <th>Body Content</th>
                    <th>File Path</th>
                </tr>
                <tr>
                    <td>Nguyen Van A</td>
                    <td class="cell-selected">anv@gmail.com</td>
                    <td>October Payslip</td>
                    <td>Dear A, your salary this month...</td>
                    <td>D:\\Salary\\Payslip_A.pdf</td>
                </tr>
                <tr>
                    <td>Tran Thi B</td>
                    <td class="cell-selected">btt@yahoo.com</td>
                    <td>Debt Notice</td>
                    <td>Dear B, you still owe...</td>
                    <td>(Leave blank if none)</td>
                </tr>
            </table>

            <h3>How to execute:</h3>
            <ol>
                <li>Prepare the data table as above.</li>
                <li><strong>IMPORTANT:</strong> Select the range containing the Email addresses (E.g., <code>B2:B3</code>). Do not select the entire table, only the Email column.</li>
                <li>Click the <strong>Send Payslip Email</strong> button on the K-Utils Ribbon.</li>
                <li>Outlook will open and create draft emails. Review them and click Send.</li>
            </ol>
        `,
        data_content: `
            <h2><i class="fas fa-database"></i> 3. Data Processing</h2>
            
            <h3>Merge Excel Files (Merge)</h3>
            <p>Consolidate data from multiple child Excel files into a single master file.</p>
            <ul>
                <li>Click the <strong>Merge Files</strong> button > Select the folder containing the files.</li>
                <li>The software will automatically open each file and copy the data into the current file.</li>
            </ul>

            <h3>Hàm RegEx (Tách dữ liệu khó)</h3>
            <p>Dùng để lấy số điện thoại hoặc Email từ chuỗi văn bản lộn xộn.</p>
            <ul>
                <li>Lấy số: <code>=RegExExtract(A1, "\\d+")</code></li>
                <li>Lấy Email: <code>=RegExExtract(A1, "[a-zA-Z0-9._-]+@[a-z]+\\.[a-z]+")</code></li>
            </ul>
        `,
        advanced_content: `
            <h2><i class="fas fa-magic"></i> 4. Tính năng Nâng cao (v2.3)</h2>
            
            <h3>Tách File (Split Workbook)</h3>
            <p>Chia nhỏ 1 sheet lớn thành nhiều file con dựa trên cột (Ví dụ: Tách lương theo Phòng ban).</p>
            <ol>
                <li>Bấm <strong>Tách File</strong>.</li>
                <li>Bước 1: Quét chọn toàn bộ bảng dữ liệu.</li>
                <li>Bước 2: Chọn 1 ô trong cột muốn tách (VD: Cột Phòng ban).</li>
                <li>Kết quả: Các file con sẽ được tạo trong thư mục mới.</li>
            </ol>

            <h3>Chuyển mã Font (TCVN3 -> Unicode)</h3>
            <div class="alert-danger">
                <strong><i class="fas fa-exclamation-circle"></i> Warning:</strong> This feature may not be 100% perfect. Please **BACK UP THE ORIGINAL FILE** before converting the font code!
            </div>
            <p>Fix messy font errors (like .VnTime) when opening old files.</p>
            <ul>
                <li>Select the region with font errors.</li>
                <li>Click <strong>Convert Font Code</strong> > Select TCVN3 (ABC) source.</li>
            </ul>
        `,
        utils_content: `
            <h2><i class="fas fa-tools"></i> 5. Tiện ích khác</h2>
            
            <div class="feature-grid">
                <div class="feature-card">
                    <i class="fas fa-qrcode feature-icon"></i>
                    <strong>Tạo QR Code</strong>
                    <p style="font-size:13px; margin-top:5px;">Chọn ô chứa link/text > Bấm nút Tạo QR. Mã QR sẽ được chèn ngay bên cạnh.</p>
                </div>
                <div class="feature-card">
                    <i class="fas fa-list feature-icon"></i>
                    <strong>Tạo Mục Lục</strong>
                    <p style="font-size:13px; margin-top:5px;">Tự động tạo sheet "Muc Luc" chứa đường dẫn (Link) đến tất cả các sheet khác trong file.</p>
                </div>
                <div class="feature-card">
                    <i class="fas fa-eye feature-icon"></i>
                    <strong>Hiện Sheet Ẩn</strong>
                    <p style="font-size:13px; margin-top:5px;">Excel mặc định chỉ cho Unhide từng sheet. Nút này giúp hiện tất cả cùng lúc.</p>
                </div>
                <div class="feature-card">
                    <i class="fas fa-calendar-alt feature-icon"></i>
                    <strong>Sửa Lỗi Ngày</strong>
                    <p style="font-size:13px; margin-top:5px;">Chuyển ngày dạng text (26.11.2023) thành dạng Date chuẩn để tính toán.</p>
                </div>
            </div>
        `,
        faq_content: `
            <h2><i class="fas fa-question-circle"></i> Frequently Asked Questions</h2>
            <ul>
                <li><strong>Q: Why don't I see the Tab after installation?</strong><br>A: See section 1 (Installation) above, follow Method 2 (Unblock file).</li>
                <li><strong>Q: Does the software contain viruses?</strong><br>A: 100% clean guarantee. This is open-source VBA code.</li>
                <li><strong>Q: Can it be used on a Macbook?</strong><br>A: No, it only supports Windows.</li>
            </ul>
        `,
        donate_content: `
            <div style="text-align: center;">
                <h2 style="border:none; justify-content: center;"><i class="fas fa-mug-hot"></i> Buy the author a cup of coffee?</h2>
                <p>The software is completely free. Your support is the motivation for me to continue development!</p>
                
                <div class="bank-card">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <span style="font-weight:800; font-size:18px;">TECHCOMBANK</span>
                        <i class="fas fa-wifi" style="transform: rotate(90deg);"></i>
                    </div>
                    <div style="margin-top:20px;">Account Number</div>
                    <div class="bank-num" id="accNum">1907 5093 4460 12</div>
                    <div style="display:flex; justify-content:space-between; align-items:flex-end;">
                        <div>
                            <div style="font-size:12px; opacity:0.8;">CARD HOLDER</div>
                            <div style="font-weight:600; font-size:16px;">TRAN TUAN KHOA</div>
                        </div>
                        <button class="copy-btn" onclick="copyToClipboard()"><i class="far fa-copy"></i></button>
                    </div>
                </div>
            </div>
        `,
        footer_content: `&copy; 2025 K-Utils PRO. Developed by Tran Tuan Khoa.`,
        toast_copied: "Account number copied!",
    },
    toast_copied: "Đã sao chép số tài khoản!",
};


// --- Hàm chuyển đổi nội dung & Ngôn ngữ ---
function updateContent(lang) {
    const data = langData[lang];
    
    // 1. Cập nhật Title và Text đơn
    document.getElementById('doc-title').innerText = data.title;
    document.getElementById('zalo-link-text').innerText = data.zaloLinkText;
    document.getElementById('toast').innerText = data.toast_copied;
    
    // 2. Cập nhật Menu Sidebar
    const navMenu = document.getElementById('nav-menu');
    navMenu.innerHTML = ''; // Xóa menu cũ
    data.nav.forEach(item => {
        if (item.group) {
            const groupDiv = document.createElement('div');
            groupDiv.className = 'nav-group';
            groupDiv.innerText = item.text;
            navMenu.appendChild(groupDiv);
        } else {
            const li = document.createElement('li');
            const link = document.createElement('a');
            link.href = item.href;
            link.className = item.accent ? 'nav-link accent' : 'nav-link';
            link.innerHTML = `<i class="${item.icon}"></i> ${item.text}`;
            link.onclick = function(e) {
                // Xử lý Highlight Section và đóng sidebar
                e.preventDefault(); 
                scrollToSection(item.href);
                if (window.innerWidth <= 768) toggleSidebar(); 
            };
            li.appendChild(link);
            navMenu.appendChild(li);
        }
    });

    // 3. Cập nhật Nội dung chính
    document.getElementById('intro-content').innerHTML = data.intro_content;
    document.getElementById('install-content').innerHTML = data.install_content;
    document.getElementById('finance-content').innerHTML = data.finance_content;
    document.getElementById('email-content').innerHTML = data.email_content;
    document.getElementById('data-content').innerHTML = data.data_content;
    document.getElementById('advanced-content').innerHTML = data.advanced_content;
    document.getElementById('utils-content').innerHTML = data.utils_content;
    document.getElementById('faq-content').innerHTML = data.faq_content;
    document.getElementById('donate-content').innerHTML = data.donate_content;
    document.getElementById('footer-content').innerHTML = data.footer_content;
    
    updateActiveLink(); 
    localStorage.setItem('kutils_lang', lang);
}

function toggleLang() {
    currentLang = currentLang === 'vi' ? 'en' : 'vi';
    updateContent(currentLang);
}

// --- Hàm chuyển đổi Theme và lưu LocalStorage ---
function toggleTheme() {
    const isDark = body.getAttribute('data-theme') === 'dark';
    const icon = document.getElementById('theme-icon');

    if (isDark) {
        body.removeAttribute('data-theme');
        icon.className = 'fas fa-moon';
        localStorage.setItem('kutils_theme', 'light');
    } else {
        body.setAttribute('data-theme', 'dark');
        icon.className = 'fas fa-sun'; 
        localStorage.setItem('kutils_theme', 'dark');
    }
}

// --- NÂNG CẤP UX: Hàm hiển thị Toast Notification ---
function showToast(message) {
    const toast = document.getElementById('toast');
    toast.innerText = message || langData[currentLang].toast_copied;
    toast.classList.add('show');
    setTimeout(function(){ toast.classList.remove('show'); }, 3000);
}


// --- NÂNG CẤP UX: Hàm sao chép và hiển thị Toast ---
function copyToClipboard() {
    const accNum = document.getElementById('accNum').innerText.replace(/\s/g, ''); 
    const btn = document.querySelector('.copy-btn');
    navigator.clipboard.writeText(accNum).then(() => {
        // 1. Hiển thị Toast
        showToast();
        // 2. Thay đổi icon trên nút
        btn.innerHTML = '<i class="fas fa-check"></i>';
        setTimeout(() => { btn.innerHTML = '<i class="far fa-copy"></i>'; }, 2000);
    });
}

function scrollToTop() {
    document.body.scrollTop = 0;
    document.documentElement.scrollTop = 0;
}

// --- NÂNG CẤP UX: Cuộn đến Section và Highlight ---
function scrollToSection(href) {
    const targetId = href.substring(1);
    const targetElement = document.getElementById(targetId);

    if (targetElement) {
        // Cuộn đến phần tử
        targetElement.scrollIntoView({ behavior: 'smooth' });

        // Thêm hiệu ứng Highlight
        targetElement.classList.remove('section-highlight'); // Reset animation
        // Cần setTimeout nhỏ để đảm bảo animation được kích hoạt lại
        setTimeout(() => {
            targetElement.classList.add('section-highlight');
        }, 50); 
    }
}

// --- Mobile Sidebar Toggle ---
function toggleSidebar() {
    sidebar.classList.toggle('open');
}

// --- Logic Active Link và Back to Top visibility (Giữ nguyên) ---
function updateActiveLink() {
    const sections = document.querySelectorAll('section');
    const navLinks = document.querySelectorAll('.nav-link');
    let current = '';
    const scrollY = window.pageYOffset || document.documentElement.scrollTop;

    sections.forEach(section => {
        const sectionTop = section.offsetTop - 150; 
        if (scrollY >= sectionTop) current = section.getAttribute('id');
    });
    
    if (scrollY < 150) current = 'intro';

    navLinks.forEach(link => {
        link.classList.remove('active');
        if (link.getAttribute('href').includes(current)) link.classList.add('active');
    });
}

window.addEventListener('scroll', () => {
    updateActiveLink();

    const backToTopBtn = document.getElementById('backToTopBtn');
    if (document.body.scrollTop > 200 || document.documentElement.scrollTop > 200) {
        backToTopBtn.style.display = "flex";
    } else {
        backToTopBtn.style.display = "none";
    }
});

// --- Khởi tạo khi tải trang (Load trạng thái từ LocalStorage) ---
document.addEventListener('DOMContentLoaded', () => {
    // 1. Load Theme
    const savedTheme = localStorage.getItem('kutils_theme');
    const icon = document.getElementById('theme-icon');
    if (savedTheme === 'dark') {
        body.setAttribute('data-theme', 'dark');
        if (icon) icon.className = 'fas fa-sun';
    } else {
        body.removeAttribute('data-theme');
        if (icon) icon.className = 'fas fa-moon';
    }
    
    // 2. Load Ngôn ngữ và Nội dung
    const savedLang = localStorage.getItem('kutils_lang') || 'vi';
    currentLang = savedLang;
    updateContent(currentLang);
});