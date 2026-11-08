/**
 * Vital Health Solution (VHS) — Invoice Generator Application Logic
 * Implements the exact fields and formulas of info_sample.xlsx
 * formatted in the visual layout of invoice_format.jpeg.
 */

(function () {
  "use strict";

  // --- State ---
  let invoiceCounter = 1;
  let lineItems = [
    { id: 1, name: "", qty: 1, unitPrice: 0, discount: 0 },
  ];
  let nextItemId = 2;

  // --- DOM Elements ---
  const elCustomerName = document.getElementById("customerName");
  const elCustomerAddress = document.getElementById("customerAddress");
  const elCustomerContact = document.getElementById("customerContact");

  const elInvoiceNo = document.getElementById("invoiceNo");
  const elInvoiceDate = document.getElementById("invoiceDate");
  const elPaymentTerms = document.getElementById("paymentTerms");
  const elCurrency = document.getElementById("currency");

  const elLineItemsContainer = document.getElementById("lineItemsContainer");
  const elBtnAddRow = document.getElementById("btnAddRow");

  const elVatRate = document.getElementById("vatRate");
  const elAmountPaid = document.getElementById("amountPaid");
  const elAutoWordsToggle = document.getElementById("autoWordsToggle");
  const elAmountInWordsInput = document.getElementById("amountInWordsInput");
  const elInvoiceNotes = document.getElementById("invoiceNotes");

  const elSellerVatBin = document.getElementById("sellerVatBin");
  const elSellerCompany = document.getElementById("sellerCompany");
  const elSellerAddress = document.getElementById("sellerAddress");
  const elSellerContact = document.getElementById("sellerContact");

  const elBtnPrint = document.getElementById("btnPrint");
  const elBtnReset = document.getElementById("btnReset");

  // --- Output Preview Elements ---
  const outSellerCompany = document.getElementById("outSellerCompany");
  const outSellerAddress = document.getElementById("outSellerAddress");
  const outSellerContact = document.getElementById("outSellerContact");
  const outSellerVatBin = document.getElementById("outSellerVatBin");

  const outCustomerName = document.getElementById("outCustomerName");
  const outCustomerAddress = document.getElementById("outCustomerAddress");
  const outCustomerContact = document.getElementById("outCustomerContact");

  const outInvoiceNo = document.getElementById("outInvoiceNo");
  const outInvoiceDate = document.getElementById("outInvoiceDate");
  const outPaymentTerms = document.getElementById("outPaymentTerms");
  const outCurrency = document.getElementById("outCurrency");

  const outPrintOn = document.getElementById("outPrintOn");
  const outItemsTbody = document.getElementById("outItemsTbody");

  const outVatRate = document.getElementById("outVatRate");
  const outSubtotal = document.getElementById("outSubtotal");
  const outVatAmount = document.getElementById("outVatAmount");
  const outGrandTotal = document.getElementById("outGrandTotal");
  const outAmountPaid = document.getElementById("outAmountPaid");
  const outBalanceDue = document.getElementById("outBalanceDue");

  const outAmountInWords = document.getElementById("outAmountInWords");
  const outNotes = document.getElementById("outNotes");

  // --- Formatting Helpers ---
  function formatMoney(num) {
    const val = Number.isFinite(num) ? num : 0;
    return val.toLocaleString("en-IN", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
  }

  function formatDateDDMMYYYY(isoDateStr) {
    if (!isoDateStr) return "—";
    const parts = isoDateStr.split("-");
    if (parts.length !== 3) return isoDateStr;
    return `${parts[2]}-${parts[1]}-${parts[0]}`;
  }

  function getTodayISO() {
    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, "0");
    const dd = String(now.getDate()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  }

  function formatPrintTimestamp() {
    const now = new Date();
    const dd = String(now.getDate()).padStart(2, "0");
    const mm = String(now.getMonth() + 1).padStart(2, "0");
    const yyyy = now.getFullYear();
    let hours = now.getHours();
    const minutes = String(now.getMinutes()).padStart(2, "0");
    const ampm = hours >= 12 ? "pm" : "am";
    hours = hours % 12;
    if (hours === 0) hours = 12;
    const hh = String(hours).padStart(2, "0");
    return `${dd}-${mm}-${yyyy} ${hh}:${minutes} ${ampm}`;
  }

  function generateInvoiceNumber(counter) {
    const year = new Date().getFullYear();
    const seq = String(counter).padStart(3, "0");
    return `INV-${year}-${seq}`;
  }

  // --- Bangladeshi Taka Number-to-Words Converter ---
  const ONES = [
    "",
    "One",
    "Two",
    "Three",
    "Four",
    "Five",
    "Six",
    "Seven",
    "Eight",
    "Nine",
    "Ten",
    "Eleven",
    "Twelve",
    "Thirteen",
    "Fourteen",
    "Fifteen",
    "Sixteen",
    "Seventeen",
    "Eighteen",
    "Nineteen",
  ];

  const TENS = [
    "",
    "",
    "Twenty",
    "Thirty",
    "Forty",
    "Fifty",
    "Sixty",
    "Seventy",
    "Eighty",
    "Ninety",
  ];

  function convertBelowThousand(n) {
    let str = "";
    if (n >= 100) {
      str += ONES[Math.floor(n / 100)] + " Hundred ";
      n %= 100;
    }
    if (n >= 20) {
      str += TENS[Math.floor(n / 10)] + " ";
      n %= 10;
    }
    if (n > 0) {
      str += ONES[n] + " ";
    }
    return str.trim();
  }

  function convertIntegerToWords(num) {
    if (num === 0) return "Zero";
    if (num < 0) return "Minus " + convertIntegerToWords(Math.abs(num));

    let words = "";
    const crore = Math.floor(num / 10000000);
    num %= 10000000;
    const lakh = Math.floor(num / 100000);
    num %= 100000;
    const thousand = Math.floor(num / 1000);
    num %= 1000;
    const remainder = num;

    if (crore > 0) {
      words += convertIntegerToWords(crore) + " Crore ";
    }
    if (lakh > 0) {
      words += convertBelowThousand(lakh) + " Lakh ";
    }
    if (thousand > 0) {
      words += convertBelowThousand(thousand) + " Thousand ";
    }
    if (remainder > 0) {
      words += convertBelowThousand(remainder);
    }

    return words.trim();
  }

  function amountToTakaWords(amount) {
    const safeAmount = Math.max(0, Number(amount) || 0);
    const totalPaisa = Math.round(safeAmount * 100);
    const taka = Math.floor(totalPaisa / 100);
    const paisa = totalPaisa % 100;

    let result = convertIntegerToWords(taka) + " Taka";
    if (paisa > 0) {
      result += " and " + convertBelowThousand(paisa) + " Paisa";
    }
    return result + " Only";
  }

  // --- Formula Calculations (Exact match to info_sample.xlsx) ---
  function computeRowTotal(item) {
    // Formula in F16: IF(B16="","",MAX(0,C16*D16-E16))
    const qty = Math.max(0, parseFloat(item.qty) || 0);
    const price = Math.max(0, parseFloat(item.unitPrice) || 0);
    const disc = Math.max(0, parseFloat(item.discount) || 0);
    return Math.max(0, qty * price - disc);
  }

  // --- Render Line Item Inputs in Editor ---
  function renderEditorLineItems() {
    elLineItemsContainer.innerHTML = "";

    lineItems.forEach((item, index) => {
      const row = document.createElement("div");
      row.className = "line-item-row";

      const slBadge = document.createElement("span");
      slBadge.className = "sl-badge";
      slBadge.textContent = index + 1;

      const nameWrap = document.createElement("div");
      nameWrap.className = "li-field li-field-name";
      nameWrap.setAttribute("data-label", "Product Name / Description");
      const nameInput = document.createElement("input");
      nameInput.type = "text";
      nameInput.placeholder = "Enter product name / description";
      nameInput.value = item.name;
      nameInput.setAttribute("aria-label", `Product ${index + 1} Name`);
      nameInput.addEventListener("input", (e) => {
        item.name = e.target.value;
        updateInvoicePreview();
      });
      nameWrap.appendChild(nameInput);

      const qtyWrap = document.createElement("div");
      qtyWrap.className = "li-field li-field-qty";
      qtyWrap.setAttribute("data-label", "Qty");
      const qtyInput = document.createElement("input");
      qtyInput.type = "number";
      qtyInput.min = "0";
      qtyInput.step = "any";
      qtyInput.placeholder = "Qty";
      qtyInput.value = item.qty;
      qtyInput.setAttribute("aria-label", `Product ${index + 1} Quantity`);
      qtyInput.addEventListener("input", (e) => {
        item.qty = e.target.value;
        rowTotalSpan.textContent = formatMoney(computeRowTotal(item));
        updateInvoicePreview();
      });
      qtyWrap.appendChild(qtyInput);

      const priceWrap = document.createElement("div");
      priceWrap.className = "li-field li-field-price";
      priceWrap.setAttribute("data-label", "Unit Price (৳)");
      const priceInput = document.createElement("input");
      priceInput.type = "number";
      priceInput.min = "0";
      priceInput.step = "0.01";
      priceInput.placeholder = "0.00";
      priceInput.value = item.unitPrice;
      priceInput.setAttribute("aria-label", `Product ${index + 1} Unit Price`);
      priceInput.addEventListener("input", (e) => {
        item.unitPrice = e.target.value;
        rowTotalSpan.textContent = formatMoney(computeRowTotal(item));
        updateInvoicePreview();
      });
      priceWrap.appendChild(priceInput);

      const discWrap = document.createElement("div");
      discWrap.className = "li-field li-field-disc";
      discWrap.setAttribute("data-label", "Discount (৳)");
      const discInput = document.createElement("input");
      discInput.type = "number";
      discInput.min = "0";
      discInput.step = "0.01";
      discInput.placeholder = "0.00";
      discInput.value = item.discount;
      discInput.setAttribute("aria-label", `Product ${index + 1} Discount`);
      discInput.addEventListener("input", (e) => {
        item.discount = e.target.value;
        rowTotalSpan.textContent = formatMoney(computeRowTotal(item));
        updateInvoicePreview();
      });
      discWrap.appendChild(discInput);

      const totalWrap = document.createElement("div");
      totalWrap.className = "li-field li-field-total";
      totalWrap.setAttribute("data-label", "Total (৳)");
      const rowTotalSpan = document.createElement("span");
      rowTotalSpan.className = "row-line-total";
      rowTotalSpan.textContent = formatMoney(computeRowTotal(item));
      totalWrap.appendChild(rowTotalSpan);

      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.className = "btn-icon-remove";
      removeBtn.title = "Remove row";
      removeBtn.innerHTML = "&times;";
      removeBtn.disabled = lineItems.length === 1;
      if (lineItems.length === 1) {
        removeBtn.style.opacity = "0.35";
        removeBtn.style.cursor = "not-allowed";
      }
      removeBtn.addEventListener("click", () => {
        if (lineItems.length > 1) {
          lineItems = lineItems.filter((r) => r.id !== item.id);
          renderEditorLineItems();
          updateInvoicePreview();
        }
      });

      row.appendChild(slBadge);
      row.appendChild(nameWrap);
      row.appendChild(qtyWrap);
      row.appendChild(priceWrap);
      row.appendChild(discWrap);
      row.appendChild(totalWrap);
      row.appendChild(removeBtn);

      elLineItemsContainer.appendChild(row);
    });
  }

  // --- Update Live Invoice Preview ---
  function updateInvoicePreview() {
    // Seller details
    outSellerCompany.textContent = elSellerCompany.value.trim() || "Vital Health Solution";
    outSellerAddress.textContent =
      elSellerAddress.value.trim() ||
      "Icon Court, House No# 14, Block-B, Avenue-2, Chandrima Model Town, Mohammadpur, Dhaka-1207, Bangladesh";
    outSellerContact.textContent =
      elSellerContact.value.trim() || "+88 01341-893323 | info@vitalhealthsolutionbd.com";
    outSellerVatBin.textContent = elSellerVatBin.value.trim() || "N/A";

    // Customer details
    outCustomerName.textContent = elCustomerName.value.trim() || "—";
    outCustomerAddress.textContent = elCustomerAddress.value.trim() || "—";
    outCustomerContact.textContent = elCustomerContact.value.trim() || "—";

    // Invoice metadata
    outInvoiceNo.textContent = elInvoiceNo.value.trim() || "—";
    outInvoiceDate.textContent = formatDateDDMMYYYY(elInvoiceDate.value);
    outPaymentTerms.textContent = elPaymentTerms.value.trim() || "Cash / Bank Transfer";
    outCurrency.textContent = elCurrency.value.trim() || "BDT (৳)";

    // Print timestamp
    outPrintOn.textContent = formatPrintTimestamp();

    // Render table rows & compute Subtotal
    outItemsTbody.innerHTML = "";
    let subtotal = 0;

    lineItems.forEach((item, index) => {
      const tr = document.createElement("tr");
      const qtyNum = Math.max(0, parseFloat(item.qty) || 0);
      const priceNum = Math.max(0, parseFloat(item.unitPrice) || 0);
      const discNum = Math.max(0, parseFloat(item.discount) || 0);
      const lineTotal = computeRowTotal(item);
      subtotal += lineTotal;

      const tdSl = document.createElement("td");
      tdSl.className = "col-sl";
      tdSl.textContent = index + 1;

      const tdDesc = document.createElement("td");
      tdDesc.className = "col-desc";
      tdDesc.textContent = item.name.trim() || "—";

      const tdQty = document.createElement("td");
      tdQty.className = "col-qty";
      tdQty.textContent = qtyNum;

      const tdPrice = document.createElement("td");
      tdPrice.className = "col-price";
      tdPrice.textContent = formatMoney(priceNum);

      const tdDisc = document.createElement("td");
      tdDisc.className = "col-disc";
      tdDisc.textContent = formatMoney(discNum);

      const tdTotal = document.createElement("td");
      tdTotal.className = "col-total";
      tdTotal.textContent = formatMoney(lineTotal);

      tr.appendChild(tdSl);
      tr.appendChild(tdDesc);
      tr.appendChild(tdQty);
      tr.appendChild(tdPrice);
      tr.appendChild(tdDisc);
      tr.appendChild(tdTotal);

      outItemsTbody.appendChild(tr);
    });

    // Summary Formulas from info_sample.xlsx:
    // F33 (Subtotal): SUM(F16:F30)
    // F34 (VAT Amount): F33 * D32 / 100
    // F35 (Grand Total): F33 + F34
    // F37 (Balance Due): MAX(0, F35 - F36)
    const vatRate = Math.max(0, parseFloat(elVatRate.value) || 0);
    const vatAmount = (subtotal * vatRate) / 100;
    const grandTotal = subtotal + vatAmount;
    const amountPaid = Math.max(0, parseFloat(elAmountPaid.value) || 0);
    const balanceDue = Math.max(0, grandTotal - amountPaid);

    outVatRate.textContent = `${vatRate}%`;
    outSubtotal.textContent = formatMoney(subtotal);
    outVatAmount.textContent = formatMoney(vatAmount);
    outGrandTotal.textContent = formatMoney(grandTotal);
    outAmountPaid.textContent = formatMoney(amountPaid);
    outBalanceDue.textContent = formatMoney(balanceDue);

    // Amount in words
    if (elAutoWordsToggle.checked) {
      const autoWords = amountToTakaWords(grandTotal);
      elAmountInWordsInput.value = autoWords;
      outAmountInWords.textContent = autoWords;
    } else {
      outAmountInWords.textContent = elAmountInWordsInput.value.trim() || "—";
    }

    // Notes
    outNotes.textContent = elInvoiceNotes.value.trim() || "—";

    updatePreviewScale();
  }

  // --- Responsive A4 Preview Scaling for Mobile & Tablet Screens ---
  function updatePreviewScale() {
    const workspace = document.getElementById("previewWorkspace");
    const wrapper = document.getElementById("invoiceScaleWrapper");
    const sheet = document.getElementById("invoiceSheet");
    if (!workspace || !wrapper || !sheet) return;

    // Constrain wrapper first so workspace measures true viewport width
    wrapper.style.width = "0px";
    wrapper.style.height = "0px";
    sheet.style.transform = "none";

    const style = window.getComputedStyle(workspace);
    const padLeft = parseFloat(style.paddingLeft) || 0;
    const padRight = parseFloat(style.paddingRight) || 0;
    const availableWidth = workspace.clientWidth - padLeft - padRight;
    const naturalWidth = sheet.offsetWidth; // ~794px (210mm)
    const naturalHeight = sheet.offsetHeight;

    if (availableWidth > 0 && naturalWidth > 0 && availableWidth < naturalWidth) {
      const scale = availableWidth / naturalWidth;
      sheet.style.transformOrigin = "top left";
      sheet.style.transform = `scale(${scale})`;
      wrapper.style.width = `${Math.floor(naturalWidth * scale)}px`;
      wrapper.style.height = `${Math.ceil(naturalHeight * scale)}px`;
    } else {
      sheet.style.transform = "none";
      wrapper.style.width = `${naturalWidth}px`;
      wrapper.style.height = "auto";
    }
  }

  // --- Event Listeners ---
  function bindEvents() {
    const inputsToWatch = [
      elCustomerName,
      elCustomerAddress,
      elCustomerContact,
      elInvoiceNo,
      elInvoiceDate,
      elPaymentTerms,
      elVatRate,
      elAmountPaid,
      elAmountInWordsInput,
      elInvoiceNotes,
      elSellerVatBin,
    ];

    inputsToWatch.forEach((inputEl) => {
      inputEl.addEventListener("input", updateInvoicePreview);
    });

    elAutoWordsToggle.addEventListener("change", () => {
      if (elAutoWordsToggle.checked) {
        elAmountInWordsInput.setAttribute("readonly", "readonly");
        elAmountInWordsInput.classList.add("input-readonly");
      } else {
        elAmountInWordsInput.removeAttribute("readonly");
        elAmountInWordsInput.classList.remove("input-readonly");
        elAmountInWordsInput.focus();
      }
      updateInvoicePreview();
    });

    elBtnAddRow.addEventListener("click", () => {
      lineItems.push({
        id: nextItemId++,
        name: "",
        qty: 1,
        unitPrice: 0,
        discount: 0,
      });
      renderEditorLineItems();
      updateInvoicePreview();

      // Focus the newly added product name input
      const rows = elLineItemsContainer.querySelectorAll(".line-item-row");
      if (rows.length > 0) {
        const lastRowInput = rows[rows.length - 1].querySelector('input[type="text"]');
        if (lastRowInput) lastRowInput.focus();
      }
    });

    elBtnPrint.addEventListener("click", () => {
      outPrintOn.textContent = formatPrintTimestamp();
      window.print();
    });

    window.addEventListener("resize", updatePreviewScale);
    window.addEventListener("beforeprint", () => {
      const wrapper = document.getElementById("invoiceScaleWrapper");
      const sheet = document.getElementById("invoiceSheet");
      if (sheet) sheet.style.transform = "none";
      if (wrapper) {
        wrapper.style.width = "100%";
        wrapper.style.height = "auto";
      }
    });
    window.addEventListener("afterprint", updatePreviewScale);

    elBtnReset.addEventListener("click", () => {
      invoiceCounter += 1;
      elCustomerName.value = "";
      elCustomerAddress.value = "";
      elCustomerContact.value = "";
      elInvoiceNo.value = generateInvoiceNumber(invoiceCounter);
      elInvoiceDate.value = getTodayISO();
      elPaymentTerms.value = "Cash / Bank Transfer";
      elVatRate.value = "0";
      elAmountPaid.value = "0";
      elAutoWordsToggle.checked = true;
      elAmountInWordsInput.setAttribute("readonly", "readonly");
      elAmountInWordsInput.classList.add("input-readonly");
      elInvoiceNotes.value =
        "Please verify product details, prices, VAT rate and applicable tax requirements before issuing.";

      lineItems = [{ id: nextItemId++, name: "", qty: 1, unitPrice: 0, discount: 0 }];
      renderEditorLineItems();
      updateInvoicePreview();
      elCustomerName.focus();
    });
  }

  // --- Initialize ---
  function init() {
    elInvoiceDate.value = getTodayISO();
    elInvoiceNo.value = generateInvoiceNumber(invoiceCounter);
    renderEditorLineItems();
    bindEvents();
    updateInvoicePreview();
  }

  init();
})();

