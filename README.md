Invoicing is an important step of ALMA Acquisition. The correct invoices help managing Acquisition data and tracking payment history. After the purchasing step, receiving physical items step or activating electronic titles step, the invoicing step jumps on the stage.
There were 3 ways to process invoices, which encompass creating invoices via EDI, creating invoices from PO/manually, and creating invoices from the Excel file. The Excel file format provided by ALMA adopts an interlacing way to store the invoice and invoice line data. By this way, you could put as more invoice lines as you like, with huge advantage. However, before Excel file was uploaded, you have to select the vendor. It probably means you have to organize your invoice Excel files vendor by vendor before importing to ALMA.
Due to the requirements to process hundreds of invoices in batch, a PHP script was written to create the Invoices, then the corresponding invoices lines data by using ALMA API. A new Excel file format was used to achieve the target. Each line in this Excel table contains 3 chunks of data which are the invoice, regular invoice line, and shipment invoice line. In very rare situations, other invoice line types will be used, so only the regular and shipment type were included.

SimpleXLSX.php was written by Sergey Shuchkin.

The MIT License (MIT)

Copyright (c) 2014 Lukas Martinelli

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
