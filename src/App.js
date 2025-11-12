import axios from "axios";
import { useState, useRef} from "react";
import * as XLSX from "xlsx";
import ReactQuill from "react-quill";
import "react-quill/dist/quill.snow.css";

function App() {
  const [subject, setSubject] = useState("");
  const [msg, setMsg] = useState("");
  const [status, setStatus] = useState(false);
  const [progress, setProgress] = useState(0);
  const [emailList, setEmailList] = useState([]);
  const [attachments, setAttachments] = useState([]);

  // Mặc định là 'TEST'
  const [accountKey, setAccountKey] = useState("TEST"); 

  const excelInputRef = useRef();
  const attachmentInputRef = useRef();
  const backendUrl =
    process.env.REACT_APP_BACKEND_URL || "http://localhost:5000";

  const handleSubject = (e) => setSubject(e.target.value);

  const processFile = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = e.target.result;
      const workbook = XLSX.read(data, { type: "binary" });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const rawList = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      const emails = rawList.map((row) => row[0]).filter(Boolean);
      setEmailList(emails);
    };
    reader.readAsBinaryString(file);
  };

  const handleFileChange = (e) => processFile(e.target.files[0]);

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    const file = e.dataTransfer.files[0];
    processFile(file);
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleAttachments = (e) => {
    const newFiles = Array.from(e.target.files);
    newFiles.forEach((file) => {
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

  const send = async () => {
    if (!subject || !msg || emailList.length === 0) {
      alert("Vui lòng nhập Chủ đề, Nội dung, và tải lên danh sách email.");
      return;
    }
    const formData = new FormData();
    formData.append("subject", subject);
    formData.append("msg", msg);
    formData.append("emailList", JSON.stringify(emailList));
    
    // Thêm accountKey vào formData
    formData.append("accountKey", accountKey);

    attachments.forEach((att) => {
      formData.append("attachments", att.file);
    });

    setStatus(true);
    setProgress(0);
    try {
      const res = await axios.post(`${backendUrl}/sendemail`, formData, {
        headers: { "Content-Type": "multipart/form-data" },
        onUploadProgress: (e) =>
          setProgress(Math.round((e.loaded * 100) / e.total)),
      });
      if (res.data.success) {
        alert("Gửi email thành công ✅");
        setSubject("");
        setMsg("");
        setEmailList([]);
        setAttachments([]);
      } else {
        alert("Gửi email thất bại ❌: " + res.data.error);
      }
    } catch (err) {
      alert("Lỗi khi gửi ❌: " + (err.response?.data?.error || err.message));
    } finally {
      setStatus(false);
      setProgress(0);
    }
  };

  const quillModules = {
    toolbar: [
      [{ header: [1, 2, 3, false] }],
      ["bold", "italic", "underline", "strike", "blockquote"],
      [{ list: "ordered" }, { list: "bullet" }],
      [{ align: [] }],
      ["link", "image"],
      ["clean"],
    ],
  };

  return (
    <div className="min-h-screen bg-gray-100 flex flex-col items-center p-6 text-gray-900">
      <header className="w-full max-w-4xl text-center py-6 text-3xl font-extrabold shadow-lg rounded-lg bg-white text-gray-800">
        Hệ thống Gửi Mail Tự động
      </header>

      <div className="bg-white shadow-2xl rounded-3xl p-8 mt-8 w-full max-w-2xl flex flex-col gap-6 transition-transform transform hover:scale-105 duration-500">
        
        {/* Dropdown với các công ty của bạn */}
        <div>
          <label htmlFor="account" className="block text-sm font-medium text-gray-700 mb-1">
            Gửi từ tài khoản:
          </label>
          <select
            id="account"
            value={accountKey}
            onChange={(e) => setAccountKey(e.target.value)}
            className="w-full p-4 border-2 border-gray-300 rounded-xl bg-white text-gray-900 focus:outline-none focus:ring-4 focus:ring-blue-500 transition duration-300"
          >
            {/* Giá trị 'value' phải khớp KEY trong .env */}
            <option value="TEST">Tài khoản Test (Cá nhân)</option>
            <option value="PVDONGYI">Công ty PVCDongyi</option>
            <option value="CIDV">Công ty CIDV</option>
            <option value="TUONGLAI">Công ty Tương Lai</option>
          </select>
        </div>
        
        <input
          type="text"
          value={subject}
          onChange={handleSubject}
          placeholder="Nhập chủ đề email"
          className="w-full p-4 border-2 border-gray-300 rounded-xl bg-white text-gray-900 placeholder-gray-500 focus:outline-none focus:ring-4 focus:ring-blue-500 transition duration-300"
        />

        <div className="quill-wrapper">
          <ReactQuill
            theme="snow"
            value={msg}
            onChange={setMsg}
            modules={quillModules}
            placeholder="Soạn nội dung email..."
          />
        </div>

        <div
          onDrop={handleDrop}
          onDragOver={handleDragOver}
          onClick={() => excelInputRef.current.click()}
          className="border-2 border-dashed border-blue-500 rounded-xl p-6 text-center cursor-pointer hover:bg-gray-50 transition duration-300 mt-10"
        >
          <p className="text-gray-600 mb-2">
            Kéo thả file Excel vào đây, hoặc nhấp để chọn
          </p>
          <p className="text-gray-500 text-sm">
            Chỉ Cột A được đọc là email
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
          <a
            href={`${backendUrl}/download-sample`}
            download="Sample_Email_List.xlsx"
            className="text-sm text-blue-600 hover:text-blue-500 hover:underline transition duration-300"
          >
            Tải xuống file Excel mẫu
          </a>
        </div>
        <p className="text-sm text-gray-600">Tổng số email: {emailList.length}</p>
        
        {emailList.length > 0 && (
          <textarea
            value={emailList.join('\n')}
            onChange={(e) => {
              const updatedList = e.target.value
                .split('\n')
                .filter(email => email.trim() !== '');
              setEmailList(updatedList);
            }}
            placeholder="Chỉnh sửa danh sách email tại đây (mỗi email một dòng)..."
            className="w-full h-40 p-4 border-2 border-gray-300 rounded-xl bg-white text-gray-900 placeholder-gray-500 focus:outline-none focus:ring-4 focus:ring-blue-500 transition duration-300 resize-none"
          />
        )}

        <div>
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
          <div className="mt-4 grid grid-cols-2 md:grid-cols-3 gap-4">
            {attachments.map((att, index) => (
              <div key={index} className="bg-gray-50 border border-gray-200 rounded-lg p-2 relative text-center">
                <button
                  onClick={() => handleRemoveAttachment(att.file.name)}
                  className="absolute -top-2 -right-2 bg-red-600 text-white rounded-full w-6 h-6 text-xs font-bold"
                >
                  X
                </button>
                {att.previewUrl ? (
                  <img 
                    src={att.previewUrl} 
                    alt="Preview" 
                    className="h-20 w-full object-contain rounded" 
                  />
                ) : (
                  <div className="h-20 flex items-center justify-center text-5xl text-gray-500">
                    📁
                  </div>
                )}
                <p className="text-xs text-gray-700 mt-2 truncate" title={att.file.name}>
                  {att.file.name}
                </p>
              </div>
            ))}
          </div>
          {attachments.length === 0 && (
             <p className="text-sm text-gray-500 mt-2">Chưa có tệp đính kèm</p>
          )}
        </div>

        {status && (
          <div className="w-full bg-gray-200 rounded-full h-4 mt-2 overflow-hidden">
            <div
              className="bg-blue-600 h-4 rounded-full transition-all duration-300"
              style={{ width: `${progress}%` }}
            ></div>
          </div>
        )}

        <button
          onClick={send}
          disabled={status}
          className={`mt-4 bg-gradient-to-r from-blue-600 via-blue-700 to-blue-800 text-white py-3 px-6 rounded-2xl font-bold text-lg shadow-lg hover:shadow-2xl transform hover:-translate-y-1 transition-all duration-300 ${
            status ? "opacity-70 cursor-not-allowed" : ""
          }`}
        >
          {status ? "Đang gửi..." : `Gửi ${emailList.length > 0 ? emailList.length : ''} Email`}
        </button>
      </div>
    </div>
  );
}

export default App;