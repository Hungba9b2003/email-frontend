// src/App.js
import axios from "axios";
import { useState, useRef } from "react";
import * as XLSX from "xlsx";
import ReactQuill from "react-quill";
import "react-quill/dist/quill.snow.css";

const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10MB

function App() {
  const [subject, setSubject] = useState("");
  const [msg, setMsg] = useState("");
  const [status, setStatus] = useState(false);
  const [progress, setProgress] = useState(0);

  // Danh sách người nhận: { email, name, gender, salutation }
  const [recipientList, setRecipientList] = useState([]);
  const [attachments, setAttachments] = useState([]);

  // Tài khoản gửi trực tiếp từ frontend
  const [smtpEmail, setSmtpEmail] = useState("");
  const [smtpPassword, setSmtpPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false); // 👈 state bật/tắt mật khẩu

  // State cho việc nhập thủ công
  const [manualEmail, setManualEmail] = useState("");
  const [manualName, setManualName] = useState("");
  const [manualGender, setManualGender] = useState("unknown"); // male, female, unknown

  const excelInputRef = useRef();
  const attachmentInputRef = useRef();

  // Có thể đổi sang env khi deploy
  const backendUrl =
    process.env.REACT_APP_BACKEND_URL || "http://localhost:5000";

  const handleSubject = (e) => setSubject(e.target.value);

  // Hàm tính toán xưng hô
  const getSalutation = (gender) => {
    if (gender === "male") return "Anh";
    if (gender === "female") return "Chị";
    return "Quý anh/chị";
  };

  // Xử lý thêm thủ công
  const handleAddManual = () => {
    if (!manualEmail) {
      alert("Vui lòng nhập ít nhất là Email!");
      return;
    }
    const newItem = {
      email: manualEmail.trim(),
      name: manualName.trim() || "",
      gender: manualGender,
      salutation: getSalutation(manualGender),
    };
    setRecipientList((prev) => [...prev, newItem]);
    // Reset form
    setManualEmail("");
    setManualName("");
    setManualGender("unknown");
  };

  // Xử lý file Excel (đọc cột Email, Tên, Giới tính)
  const processFile = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = e.target.result;
      const workbook = XLSX.read(data, { type: "binary" });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      const rawList = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      const newRecipients = rawList
        .slice(1) // Bỏ dòng tiêu đề
        .map((row) => {
          const email = row[0]?.toString().trim();
          const name = (row[1] || "").toString().trim();
          const genderRaw = row[2] ? row[2].toString().toLowerCase() : "";

          if (!email) return null;

          let gender = "unknown";
          if (genderRaw.includes("nam") || genderRaw.includes("male"))
            gender = "male";
          if (
            genderRaw.includes("nữ") ||
            genderRaw.includes("nu") ||
            genderRaw.includes("female")
          )
            gender = "female";

          return {
            email,
            name,
            gender,
            salutation: getSalutation(gender),
          };
        })
        .filter(Boolean);

      setRecipientList((prev) => [...prev, ...newRecipients]);
    };
    reader.readAsBinaryString(file);
  };

  const handleFileChange = (e) => processFile(e.target.files[0]);

  const handleRemoveRecipient = (index) => {
    setRecipientList((prev) => prev.filter((_, i) => i !== index));
  };

  const handleAttachments = (e) => {
    const newFiles = Array.from(e.target.files);

    newFiles.forEach((file) => {
      if (file.size > MAX_FILE_SIZE) {
        alert(`File "${file.name}" quá lớn. Vui lòng chọn file <= 10MB.`);
        return;
      }

      let previewUrl = null;
      if (file.type.startsWith("image/")) {
        previewUrl = URL.createObjectURL(file);
      }
      setAttachments((prev) => [...prev, { file, previewUrl }]);
    });

    e.target.value = null;
  };

  const handleRemoveAttachment = (fileName) => {
    setAttachments((prev) => {
      const fileToRemove = prev.find((att) => att.file.name === fileName);
      if (fileToRemove && fileToRemove.previewUrl) {
        URL.revokeObjectURL(fileToRemove.previewUrl);
      }
      return prev.filter((att) => att.file.name !== fileName);
    });
  };

  const downloadSample = () => {
    const data = [
      ["Email", "Tên", "Giới tính"], // Header
      ["nguyenva@company.com", "Nguyễn Văn A", "Nam"],
      ["tranthib@company.com", "Trần Thị B", "Nữ"],
      ["partner@company.com", "Đối Tác C", "Other"],
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "DS_Mau");

    XLSX.writeFile(workbook, "Mau_Danh_Sach_Email.xlsx");
  };

  const send = async () => {
    if (!smtpEmail || !smtpPassword) {
      alert("Vui lòng nhập Email gửi và Mật khẩu/App Password.");
      return;
    }

    if (!subject || !msg || recipientList.length === 0) {
      alert("Vui lòng nhập Chủ đề, Nội dung và danh sách người nhận.");
      return;
    }

    console.log("📦 Data gửi đi:", {
      subject,
      msg,
      recipientList,
      smtpEmail,
      attachments: attachments.map((a) => a.file.name),
    });

    const formData = new FormData();
    formData.append("subject", subject);
    formData.append("msg", msg);
    formData.append("recipientList", JSON.stringify(recipientList));
    formData.append("smtpEmail", smtpEmail);
    formData.append("smtpPassword", smtpPassword);

    attachments.forEach((att) => {
      formData.append("attachments", att.file);
    });

    setStatus(true);
    setProgress(0);
    try {
      const res = await axios.post(`${backendUrl}/sendemail`, formData, {
        // KHÔNG set Content-Type, để browser tự thêm boundary
        onUploadProgress: (e) => {
          if (e.total) {
            setProgress(Math.round((e.loaded * 100) / e.total));
          }
        },
      });

      console.log("📩 Phản hồi server:", res.data);

      if (res.data.success) {
        alert(res.data.message || "Gửi email thành công ✅");
        setSubject("");
        setMsg("");
        setRecipientList([]);
        setAttachments([]);
        // Không xoá smtpEmail/smtpPassword để user gửi tiếp nếu muốn
      } else {
        alert("Gửi email thất bại ❌: " + (res.data.error || "Không rõ lỗi"));
      }
    } catch (err) {
      console.error("❌ Lỗi khi gửi:", err);
      alert("Lỗi khi gửi ❌: " + (err.response?.data?.error || err.message));
    } finally {
      setStatus(false);
      setProgress(0);
    }
  };

  const quillModules = {
    toolbar: [
      [{ header: [1, 2, 3, false] }],
      ["bold", "italic", "underline", "strike"],
      [{ list: "ordered" }, { list: "bullet" }],
      [{ align: [] }],
      ["link", "image", "clean"],
    ],
  };

  return (
    <div className="min-h-screen bg-gray-100 flex flex-col items-center p-6 text-gray-900">
      <header className="w-full max-w-5xl text-center py-6 text-3xl font-extrabold shadow-lg rounded-lg bg-white text-gray-800">
        Hệ thống Gửi Mail Tự động 🚀
      </header>

      <div className="bg-white shadow-2xl rounded-3xl p-8 mt-8 w-full max-w-5xl flex flex-col gap-6">
        {/* Gửi từ & Chủ đề */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div>
            <label className="block text-sm font-bold text-gray-700 mb-1">
              Email gửi (Gmail):
            </label>
            <input
              type="email"
              value={smtpEmail}
              onChange={(e) => setSmtpEmail(e.target.value)}
              placeholder="vd: yourmail@gmail.com"
              className="w-full p-3 border rounded-xl bg-gray-50 focus:ring-2 focus:ring-blue-500 mb-2"
            />
            <label className="block text-sm font-bold text-gray-700 mb-1">
              Mật khẩu / App Password:
            </label>
            {/* Ô mật khẩu có nút 👁️ bật/tắt */}
            <div className="relative">
              <input
                type={showPassword ? "text" : "password"}
                value={smtpPassword}
                onChange={(e) => setSmtpPassword(e.target.value)}
                placeholder="Mật khẩu ứng dụng Gmail"
                className="w-full p-3 border rounded-xl bg-gray-50 focus:ring-2 focus:ring-blue-500 pr-10"
              />
              <button
                type="button"
                onClick={() => setShowPassword((prev) => !prev)}
                className="absolute inset-y-0 right-0 px-3 flex items-center text-gray-500 hover:text-gray-700"
              >
                {showPassword ? "🙈" : "👁️"}
              </button>
            </div>
            <p className="text-xs text-gray-500 mt-1">
              ⚠️ Nên dùng <b>App Password</b> của Gmail, không dùng mật khẩu
              chính.
            </p>
          </div>
          <div>
            <label className="block text-sm font-bold text-gray-700 mb-1">
              Chủ đề:
            </label>
            <input
              type="text"
              value={subject}
              onChange={handleSubject}
              placeholder="Nhập chủ đề email"
              className="w-full p-3 border rounded-xl bg-gray-50 focus:ring-2 focus:ring-blue-500"
            />
          </div>
        </div>

        {/* Soạn thảo nội dung */}
        <div className="quill-wrapper">
          <p className="text-xs text-gray-500 mb-1">
            Mẹo: Dùng <b>{`{{danh_xung}}`}</b> để hiện Anh/Chị,{" "}
            <b>{`{{ten}}`}</b> để hiện tên người nhận.
          </p>
          <ReactQuill
            theme="snow"
            value={msg}
            onChange={setMsg}
            modules={quillModules}
            placeholder="Soạn nội dung email..."
            className="bg-white"
          />
        </div>

        <hr className="border-gray-200" />

        {/* Khu vực quản lý danh sách người nhận */}
        <div className="flex flex-col gap-4">
          <h3 className="text-xl font-bold text-gray-800">
            Danh sách người nhận ({recipientList.length})
          </h3>

          {/* 1. Form thêm thủ công */}
          <div className="bg-blue-50 p-4 rounded-xl border border-blue-100 grid grid-cols-1 md:grid-cols-4 gap-3 items-end">
            <div>
              <label className="text-xs font-semibold text-gray-600">
                Email (*)
              </label>
              <input
                type="email"
                value={manualEmail}
                onChange={(e) => setManualEmail(e.target.value)}
                placeholder="example@gmail.com"
                className="w-full p-2 border rounded-lg text-sm"
              />
            </div>
            <div>
              <label className="text-xs font-semibold text-gray-600">
                Tên
              </label>
              <input
                type="text"
                value={manualName}
                onChange={(e) => setManualName(e.target.value)}
                placeholder="Nguyễn Văn A"
                className="w-full p-2 border rounded-lg text-sm"
              />
            </div>
            <div>
              <label className="text-xs font-semibold text-gray-600">
                Giới tính
              </label>
              <select
                value={manualGender}
                onChange={(e) => setManualGender(e.target.value)}
                className="w-full p-2 border rounded-lg text-sm"
              >
                <option value="unknown">Không rõ</option>
                <option value="male">Nam (Anh)</option>
                <option value="female">Nữ (Chị)</option>
              </select>
            </div>
            <button
              onClick={handleAddManual}
              className="bg-blue-600 text-white p-2 rounded-lg font-semibold hover:bg-blue-700 transition text-sm h-[38px]"
            >
              + Thêm Lẻ
            </button>
          </div>

          {/* 2. Import Excel */}
          <div
            onClick={() => excelInputRef.current.click()}
            className="border-2 border-dashed border-gray-300 rounded-xl p-4 text-center cursor-pointer hover:bg-gray-50 transition text-sm text-gray-500"
          >
            <p>
              📂 Bấm để nhập thêm từ Excel (Cột A: Email, Cột B: Tên, Cột C:
              Nam/Nữ)
            </p>
            <input
              ref={excelInputRef}
              type="file"
              accept=".xlsx, .xls"
              onChange={handleFileChange}
              className="hidden"
            />
          </div>

          <div className="text-center mt-2">
            <button
              onClick={downloadSample}
              className="text-sm text-blue-600 hover:text-blue-800 font-semibold hover:underline transition duration-300 flex items-center justify-center gap-1 mx-auto"
            >
              📥 Tải xuống file Excel mẫu (Chuẩn)
            </button>
          </div>

          {/* 3. Bảng hiển thị danh sách */}
          {recipientList.length > 0 ? (
            <div className="overflow-x-auto max-h-60 overflow-y-auto border border-gray-200 rounded-lg">
              <table className="min-w-full divide-y divide-gray-200">
                <thead className="bg-gray-100 sticky top-0">
                  <tr>
                    <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">
                      #
                    </th>
                    <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">
                      Email
                    </th>
                    <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">
                      Tên
                    </th>
                    <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">
                      Xưng hô
                    </th>
                    <th className="px-4 py-2 text-center text-xs font-medium text-gray-500 uppercase">
                      Xóa
                    </th>
                  </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-200">
                  {recipientList.map((item, idx) => (
                    <tr key={idx} className="hover:bg-gray-50">
                      <td className="px-4 py-2 text-sm text-gray-500">
                        {idx + 1}
                      </td>
                      <td className="px-4 py-2 text-sm font-medium text-gray-900">
                        {item.email}
                      </td>
                      <td className="px-4 py-2 text-sm text-gray-500">
                        {item.name || "-"}
                      </td>
                      <td className="px-4 py-2 text-sm text-blue-600 font-medium">
                        {item.salutation} {item.name}
                      </td>
                      <td className="px-4 py-2 text-center">
                        <button
                          onClick={() => handleRemoveRecipient(idx)}
                          className="text-red-500 hover:text-red-700 font-bold"
                        >
                          ✕
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <p className="text-center text-gray-400 italic text-sm">
              Chưa có người nhận nào trong danh sách.
            </p>
          )}
        </div>

        {/* Attachment & Send Button */}
        <div className="mt-4">
          <label
            htmlFor="attachments"
            className="font-medium text-blue-600 cursor-pointer hover:text-blue-500"
          >
            📎 Thêm tệp đính kèm
          </label>
          <input
            id="attachments"
            ref={attachmentInputRef}
            type="file"
            multiple
            onChange={handleAttachments}
            className="hidden"
          />

          <div className="mt-2 flex flex-wrap gap-2">
            {attachments.map((att, index) => (
              <span
                key={index}
                className="bg-gray-100 text-xs px-2 py-1 rounded border flex items-center gap-2"
              >
                {att.file.name}
                <button
                  onClick={() => handleRemoveAttachment(att.file.name)}
                  className="text-red-500 font-bold"
                >
                  x
                </button>
              </span>
            ))}
          </div>
        </div>

        {status && (
          <div className="w-full bg-gray-200 rounded-full h-4 overflow-hidden">
            <div
              className="bg-blue-600 h-4 rounded-full transition-all duration-300"
              style={{ width: `${progress}%` }}
            ></div>
          </div>
        )}

        <button
          onClick={send}
          disabled={status}
          className={`w-full bg-gradient-to-r from-blue-600 to-blue-800 text-white py-3 rounded-2xl font-bold text-lg shadow-lg hover:shadow-xl transition-all ${
            status ? "opacity-70 cursor-not-allowed" : ""
          }`}
        >
          {status ? "Đang gửi..." : `Gửi ${recipientList.length} Email`}
        </button>
      </div>
    </div>
  );
}

export default App;
