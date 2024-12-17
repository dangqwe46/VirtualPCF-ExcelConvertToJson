import * as React from "react";
import { IIconProps, PrimaryButton } from "@fluentui/react";
import * as XLSX from "xlsx";
import "./css/ExcelConvertJSON.css";

export interface IFile {
  name: string;
  file: string;
}
export interface IFileUploaderProps {
  stateChanged: () => void;
  files: (files: IFile[]) => void;
  jsonOutput: (jsonOutput: string) => void;
  label: string | null;
  accepts: string | null;
  icon: string | null;
  resetFiles: string | null;
}

export const FileUploader = (props: IFileUploaderProps) => {
  const inputRef = React.useRef<HTMLInputElement>(null);
  const [files, setFiles] = React.useState<IFile[]>([]);
  const [jsonOutput, setJsonOutput] = React.useState<string>("");
  const { label, accepts, icon, resetFiles } = props;

  const triggerUpload = React.useCallback(() => {
    if (inputRef && inputRef.current) {
      inputRef.current.click();
    }
  }, []);

  React.useEffect(() => {
    setFiles([]);
  }, [resetFiles]);

  React.useEffect(() => {
    setJsonOutput("");
  }, [resetFiles]);

  React.useEffect(() => {
    props.files(files);
    props.stateChanged();
  }, [files]);
  React.useEffect(() => {
    props.jsonOutput(jsonOutput);
    props.stateChanged();
  }, [jsonOutput]);

  const readFiles = React.useCallback(
    (arrayFiles: File[]) => {
      const fileArray: IFile[] = [];

      arrayFiles.map(async (file) => {
        const fileReader = new FileReader();
        const fileReader2 = new FileReader();
        fileReader.onload = (e) => {
          // eslint-disable-next-line no-debugger
          debugger;
          const data = new Uint8Array(e.target!.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: "array" });
          const jsonOutput = convertExcelToJson(workbook);
          setJsonOutput(jsonOutput);
        };
         fileReader.readAsArrayBuffer(file);
        fileReader2.onloadend = () => {
          fileArray.push({
            name: file.name,
            file: fileReader2.result as string,
          });
          setFiles([...files, ...fileArray]);
        };
        fileReader2.readAsDataURL(file);
      });
    },
    [files]
  );

  const fileChanged = React.useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      if (e.target.files) {
        const arrayFiles = Array.from(e.target.files);
        readFiles(arrayFiles);
      }
    },
    [files]
  );
  const convertExcelToJson = function (workbook: XLSX.WorkBook): string {
    const jsonOutput: { [sheetName: string]: string[] } = {};
    workbook.SheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      jsonOutput[sheetName] = XLSX.utils.sheet_to_json(worksheet);
    });
    return JSON.stringify(jsonOutput);
  };
  const actionIconObject: IIconProps = {
    iconName: icon ? icon : "Attach",
  };

  return (
    <>
      <PrimaryButton iconProps={actionIconObject} onClick={triggerUpload}>
        {label}
      </PrimaryButton>
      <input
        type="file"
        value=""
        multiple={false}
        ref={inputRef}
        accept={accepts ? accepts : ""}
        onChange={fileChanged}
        style={{
          display: "none",
        }}
      />
    </>
  );
};
