import { useState } from "react";
import "./App.css";

function App() {

  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);

  const handleUpload = async () => {

    if (!file) {
      alert("Please upload a resume");
      return;
    }

    setLoading(true);

    try {

      const formData = new FormData();
      formData.append("file", file);

      const response = await fetch(
        "http://localhost:5000/generate-ppt",
        {
          method: "POST",
          body: formData,
        }
      );

      if (!response.ok) {
        throw new Error("Server failed");
      }

      const blob = await response.blob();

      const url = window.URL.createObjectURL(blob);

      const link = document.createElement("a");

      link.href = url;

      // =========================
      // GET FILENAME FROM BACKEND
      // =========================

      const contentDisposition =
        response.headers.get("Content-Disposition");

      let filename = "Resume_Presentation.pptx";

      if (contentDisposition) {

        const match = contentDisposition.match(
          /filename="?([^"]+)"?/
        );

        if (match?.[1]) {
          filename = match[1];
        }
      }

      link.download = filename;

      document.body.appendChild(link);

      link.click();

      document.body.removeChild(link);

      setTimeout(() => {
        window.URL.revokeObjectURL(url);
      }, 100);

      alert(
        "✅ Successful Download! PPT saved to Downloads folder."
      );

    } catch (error) {

      console.error(
        "Non-critical frontend warning:",
        error
      );

    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="page">

      <div className="hero">

        <div className="badge">
          ✨ AI-powered with Groq
        </div>

        <h1>
          Resume → Beautiful PowerPoint
        </h1>

        <p>
          Upload your PDF or DOCX resume.
          We'll structure it with AI and generate
          a polished presentation you can
          download instantly.
        </p>

      </div>

      <div className="upload-card">

        <div className="upload-icon">
          ⬆️
        </div>

        <h2>
          Upload your resume
        </h2>

        <p className="subtext">
          Drag & drop a PDF or DOCX file,
          or click to browse
        </p>

        <input
          type="file"
          accept=".pdf,.docx"
          id="fileUpload"
          hidden
          onChange={(e) =>
            setFile(e.target.files[0])
          }
        />

        <label
          htmlFor="fileUpload"
          className="file-btn"
        >
          📄 Choose file
        </label>

        {file && (
          <p className="file-name">
            {file.name}
          </p>
        )}

        <button
          className="generate-btn"
          onClick={handleUpload}
          disabled={loading}
        >
          {loading
            ? "Generating..."
            : "Generate PPT"}
        </button>

      </div>

    </div>
  );
}

export default App;