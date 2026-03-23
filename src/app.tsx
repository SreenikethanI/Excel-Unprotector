import "./app.css";
import { XLSX_MIME_TYPE } from "./constants.ts";
import { unprotectWorkbook } from "./logic.ts";

export function App() {
  async function handleFileInput(event: Event) {
    const files = (event.target as HTMLInputElement).files;
    if (!files || files.length != 1) {
      return;
    }
    const file = files[0];
    const result = await unprotectWorkbook(file);
    console.log(result);
    console.log("", window.URL.createObjectURL(result));
  }

  return (
    <>
      <input type="file" accept={XLSX_MIME_TYPE} onChange={handleFileInput} />
    </>
  );
}
