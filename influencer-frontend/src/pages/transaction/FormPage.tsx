import * as XLSX from "xlsx";

import {
  Alert,
  Box,
  Button,
  Card,
  CardContent,
  Container,
  Snackbar,
  Stack,
  Typography,
} from "@mui/material";
import React, { useState } from "react";
import {
  downloadInvoicePDF,
  downloadInvoicesAsZip,
} from "../../apis/invoiceApi";

import { styled } from "@mui/material/styles";
import jsPDF from "jspdf";
import { useForm } from "react-hook-form";

interface BankTransferInfo {
  invoiceNum: string;
  amount: number;
  firstName: string;
  lastName: string;
  bankName: string;
  bankBranch: string;
  accountNumber: string;
  city: string;
  address: string;
  fullName?: string;
  fullAddress?: string;
}

const VisuallyHiddenInput = styled("input")`
  clip: rect(0 0 0 0);
  clip-path: inset(50%);
  height: 1px;
  overflow: hidden;
  position: absolute;
  bottom: 0;
  left: 0;
  white-space: nowrap;
  width: 1px;
`;

const FormPage: React.FC = () => {
  const [transferInfo, setTransferInfo] = useState<BankTransferInfo[]>([]);
  const [currentIndex, setCurrentIndex] = useState(0);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);
  const { handleSubmit } = useForm();

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    try {
      setIsLoading(true);
      setError(null);
      const file = e.target.files?.[0];
      if (!file) {
        setError("파일을 선택해주세요.");
        return;
      }

      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const data = new Uint8Array(evt.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: "array" });

          // Check if '입력' sheet exists
          if (!workbook.SheetNames.includes("입력")) {
            throw new Error("'입력' 시트를 찾을 수 없습니다.");
          }

          const sheet = workbook.Sheets["입력"];

          console.log("Available sheets:", workbook.SheetNames);
          console.log("Selected sheet:", "입력");

          const rawDataArray = XLSX.utils.sheet_to_json(sheet, {
            header: "A",
          });

          console.log("Raw Excel data:", rawDataArray);
          console.log("Total rows found:", rawDataArray.length);

          // Skip first two rows and process from third row
          const processedData = rawDataArray
            .slice(2)
            .map((row: any, index: number) => {
              console.log(`Processing row ${index + 3}:`, row);

              // Log specific columns we're interested in
              console.log(`Row ${index + 3} relevant columns:`, {
                invoiceNum: row["A"],
                amount: row["O"],
                firstName: row["P"],
                lastName: row["R"],
                bankName: row["T"],
                bankBranch: row["W"],
                accountNumber: row["Z"],
                city: row["AC"],
                address: row["AD"],
              });

              // Convert amount to number and handle potential formatting
              const amount =
                typeof row["O"] === "string"
                  ? Number(row["O"].replace(/[^0-9.-]+/g, ""))
                  : Number(row["O"]) || 0;

              if (!row["P"] || !row["R"] || !row["T"]) {
                console.warn(`Row ${index + 3} missing required data:`, {
                  firstName: row["P"],
                  lastName: row["R"],
                  bankName: row["T"],
                });
                return null; // Return null instead of throwing error
              }

              const info: BankTransferInfo = {
                invoiceNum: row["A"] || "",
                amount: amount,
                firstName: row["P"] || "",
                lastName: row["R"] || "",
                bankName: row["T"] || "",
                bankBranch: row["W"] || "",
                accountNumber: row["Z"] || "",
                city: row["AC"] || "",
                address: row["AD"] || "",
                fullName: `${row["R"] || ""} ${row["P"] || ""}`.trim(),
                fullAddress: `${row["AC"] || ""} ${row["AD"] || ""}`.trim(),
              };

              console.log(`Processed row ${index + 3}:`, info);
              return info;
            })
            .filter((item): item is BankTransferInfo => {
              // Check if item is not null and has all required fields with non-empty values
              const isValid =
                item !== null &&
                typeof item.firstName === "string" &&
                item.firstName.length > 0 &&
                typeof item.lastName === "string" &&
                item.lastName.length > 0 &&
                typeof item.bankName === "string" &&
                item.bankName.length > 0;

              if (!isValid) {
                console.warn("Filtered out invalid item:", item);
              }

              return isValid;
            });

          if (processedData.length === 0) {
            console.error("No valid data found after processing");
            throw new Error("처리할 데이터가 없습니다.");
          }

          console.log("Final processed data summary:", {
            totalRows: rawDataArray.length,
            validRows: processedData.length,
            totalAmount: processedData.reduce(
              (sum, item) => sum + item.amount,
              0
            ),
            data: processedData,
          });
          setTransferInfo(processedData);
          setSuccess(true);
        } catch (err) {
          setError(
            err instanceof Error
              ? err.message
              : "파일 처리 중 오류가 발생했습니다."
          );
        } finally {
          setIsLoading(false);
        }
      };

      reader.onerror = () => {
        setError("파일 읽기 중 오류가 발생했습니다.");
        setIsLoading(false);
      };

      reader.readAsArrayBuffer(file);
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "알 수 없는 오류가 발생했습니다."
      );
      setIsLoading(false);
    }
  };

  const formatAmount = (amount: number) => {
    return amount.toLocaleString("ja-JP");
  };

  const handleCloseSnackbar = () => {
    setSuccess(false);
    setError(null);
  };

  const handleNext = () => {
    if (currentIndex < transferInfo.length - 1) {
      setCurrentIndex(currentIndex + 1);
    }
  };

  const handlePrevious = () => {
    if (currentIndex > 0) {
      setCurrentIndex(currentIndex - 1);
    }
  };

  const onSubmit = async () => {
    try {
      setIsLoading(true);
      setError(null);

      if (transferInfo.length === 0) {
        throw new Error("제출할 데이터가 없습니다.");
      }

      // Create PDF document
      const doc = new jsPDF();
      const pageWidth = doc.internal.pageSize.getWidth();

      // Add title
      doc.setFontSize(20);
      doc.text("Bank Transfer Information", pageWidth / 2, 20, {
        align: "center",
      });

      // Add date
      doc.setFontSize(12);
      const currentDate = new Date().toLocaleDateString("ko-KR");
      doc.text(`Generated on: ${currentDate}`, pageWidth / 2, 30, {
        align: "center",
      });

      // Add content
      doc.setFontSize(10);
      let yPosition = 50;
      const lineHeight = 7;

      transferInfo.forEach((info, index) => {
        if (yPosition > 250) {
          // Check if we need a new page
          doc.addPage();
          yPosition = 20;
        }

        doc.setFontSize(12);
        doc.text(`Transfer #${index + 1}`, 20, yPosition);
        yPosition += lineHeight;

        doc.setFontSize(10);
        doc.text(`Name: ${info.fullName}`, 25, yPosition);
        yPosition += lineHeight;
        doc.text(`Invoice Number: ${info.invoiceNum}`, 25, yPosition);
        yPosition += lineHeight;
        doc.text(`Amount: ¥${formatAmount(info.amount)}`, 25, yPosition);
        yPosition += lineHeight;
        doc.text(`Bank: ${info.bankName}`, 25, yPosition);
        yPosition += lineHeight;
        doc.text(`Branch: ${info.bankBranch}`, 25, yPosition);
        yPosition += lineHeight;
        doc.text(`Account: ${info.accountNumber}`, 25, yPosition);
        yPosition += lineHeight;
        doc.text(`Address: ${info.fullAddress}`, 25, yPosition);
        yPosition += lineHeight * 2; // Add extra space between entries
      });

      // Add summary at the end
      doc.addPage();
      doc.setFontSize(14);
      doc.text("Summary", pageWidth / 2, 20, { align: "center" });
      doc.setFontSize(12);
      const totalAmount = transferInfo.reduce(
        (sum, item) => sum + item.amount,
        0
      );
      doc.text(`Total Transfers: ${transferInfo.length}`, 20, 40);
      doc.text(`Total Amount: ¥${formatAmount(totalAmount)}`, 20, 50);

      // Save the PDF
      const fileName = `bank_transfers_${currentDate.replace(/\//g, "-")}.pdf`;
      doc.save(fileName);

      setSuccess(true);
      return true;
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "제출 중 오류가 발생했습니다."
      );
      throw err;
    } finally {
      setIsLoading(false);
    }
  };

  const handleDownloadInvoice = async () => {
    try {
      const currentInfo = transferInfo[currentIndex];
      await downloadInvoicePDF({
        fullName: currentInfo.fullName || "",
        fullAddress: currentInfo.fullAddress || "",
        amount: currentInfo.amount,
        bankName: currentInfo.bankName,
        bankBranch: currentInfo.bankBranch,
        accountNumber: currentInfo.accountNumber,
        invoiceNum: currentInfo.invoiceNum,
      });
      setSuccess(true);
    } catch (err) {
      setError("인보이스 다운로드에 실패했습니다.");
      console.error(err);
    }
  };

  const handleSaveAll = async () => {
    try {
      setIsLoading(true);
      // Transform the data to match the API's expected format
      const allTransferInfo = transferInfo.map((info) => ({
        fullName: `${info.lastName} ${info.firstName}`.trim(),
        fullAddress: `${info.city} ${info.address}`.trim(),
        amount: Number(info.amount),
        bankName: info.bankName,
        bankBranch: info.bankBranch,
        accountNumber: info.accountNumber,
        invoiceNum: info.invoiceNum,
      }));

      await downloadInvoicesAsZip(allTransferInfo);
      setSuccess(true);
    } catch (error) {
      setError("인보이스 다운로드 중 오류가 발생했습니다.");
      console.error(error);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <Container maxWidth="lg" sx={{ py: 4 }}>
      <Typography variant="h4" component="h1" gutterBottom>
        송금 정보 양식
      </Typography>

      <Box sx={{ mb: 4 }}>
        <Button
          component="label"
          variant="contained"
          sx={{ mb: 2 }}
          disabled={isLoading}
        >
          Excel 파일 업로드
          <VisuallyHiddenInput
            type="file"
            accept=".xlsx, .xls"
            onChange={handleFileUpload}
          />
        </Button>
      </Box>

      <Box component="form" onSubmit={handleSubmit(onSubmit)}>
        {transferInfo.length > 0 && (
          <>
            <Box
              sx={{
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                gap: 2,
                mb: 4,
              }}
            >
              <Button
                onClick={handlePrevious}
                disabled={currentIndex === 0}
                variant="contained"
                color="primary"
                sx={{ minWidth: "120px" }}
              >
                이전
              </Button>
              <Typography>
                {currentIndex + 1} / {transferInfo.length}
              </Typography>
              <Button
                onClick={handleNext}
                disabled={currentIndex === transferInfo.length - 1}
                variant="contained"
                color="primary"
                sx={{ minWidth: "120px" }}
              >
                다음
              </Button>
            </Box>

            <Card variant="outlined" sx={{ maxWidth: 600, mx: "auto" }}>
              <CardContent>
                <Typography variant="h6" gutterBottom>
                  {transferInfo[currentIndex].fullName}
                </Typography>
                <Stack spacing={1}>
                  <Box>
                    <Typography component="span" fontWeight="medium">
                      Invoice Number:{" "}
                    </Typography>
                    <Typography component="span">
                      {transferInfo[currentIndex].invoiceNum}
                    </Typography>
                  </Box>
                  <Box>
                    <Typography component="span" fontWeight="medium">
                      Amount:{" "}
                    </Typography>
                    <Typography component="span">
                      ¥{formatAmount(transferInfo[currentIndex].amount)}
                    </Typography>
                  </Box>
                  <Box>
                    <Typography component="span" fontWeight="medium">
                      Bank Name:{" "}
                    </Typography>
                    <Typography component="span">
                      {transferInfo[currentIndex].bankName}
                    </Typography>
                  </Box>
                  <Box>
                    <Typography component="span" fontWeight="medium">
                      Bank Branch:{" "}
                    </Typography>
                    <Typography component="span">
                      {transferInfo[currentIndex].bankBranch}
                    </Typography>
                  </Box>
                  <Box>
                    <Typography component="span" fontWeight="medium">
                      Account Number:{" "}
                    </Typography>
                    <Typography component="span">
                      {transferInfo[currentIndex].accountNumber}
                    </Typography>
                  </Box>
                  <Box>
                    <Typography component="span" fontWeight="medium">
                      Full Address:{" "}
                    </Typography>
                    <Typography component="span">
                      {transferInfo[currentIndex].fullAddress}
                    </Typography>
                  </Box>
                </Stack>
                <Box sx={{ mt: 2, display: "flex", justifyContent: "center" }}>
                  <Button
                    onClick={handleDownloadInvoice}
                    variant="contained"
                    color="secondary"
                    sx={{ mt: 2 }}
                    disabled={isLoading}
                  >
                    {isLoading ? "저장중..." : "인보이스 다운로드"}
                  </Button>
                </Box>
              </CardContent>
            </Card>

            <Box sx={{ display: "flex", justifyContent: "center", mt: 4 }}>
              <Button
                variant="contained"
                color="primary"
                onClick={handleSaveAll}
                disabled={isLoading || transferInfo.length === 0}
              >
                {isLoading ? "저장중..." : "저장 (모두)"}
              </Button>
            </Box>
          </>
        )}
      </Box>

      <Snackbar
        open={!!error || success}
        autoHideDuration={6000}
        onClose={handleCloseSnackbar}
      >
        <Alert
          onClose={handleCloseSnackbar}
          severity={error ? "error" : "success"}
          sx={{ width: "100%" }}
        >
          {error || "파일이 성공적으로 처리되었습니다."}
        </Alert>
      </Snackbar>
    </Container>
  );
};

export default FormPage;
