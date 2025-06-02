export interface BankTransferInfo {
  invoiceNum: string;
  fullName: string;
  fullAddress: string;
  amount: number;
  bankName: string;
  bankBranch: string;
  accountNumber: string;
}

export const downloadInvoicePDF = async (info: BankTransferInfo) => {
  try {
    const response = await fetch('http://localhost:3000/invoice/generate', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(info),
    });

    if (!response.ok) {
      throw new Error('Failed to generate invoice');
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${info.invoiceNum}.pdf`;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);
  } catch (error) {
    console.error('Error downloading invoice:', error);
    throw error;
  }
};

export const downloadInvoicesAsZip = async (invoicesData: BankTransferInfo[]) => {
  try {
    const response = await fetch(`http://localhost:3000/invoice/generate-zip`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/zip'
      },
      body: JSON.stringify({ invoices: invoicesData }),
    });

    if (!response.ok) {
      throw new Error('Failed to generate invoices zip');
    }

    const blob = await response.blob();
    
    // Get filename from Content-Disposition header or use default
    const contentDisposition = response.headers.get('Content-Disposition');
    let filename = 'invoice.zip';
    if (contentDisposition) {
      const matches = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/.exec(contentDisposition);
      if (matches != null && matches[1]) {
        filename = matches[1].replace(/['"]/g, '');
      }
    }

    // Create download link
    const url = window.URL.createObjectURL(new Blob([blob], { type: 'application/zip' }));
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);
  } catch (error) {
    console.error('Error downloading invoices zip:', error);
    throw error;
  }
};
