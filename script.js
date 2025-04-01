

document.addEventListener("DOMContentLoaded", function () {
  function saveTableData() {
    const tableBody = document.getElementById("repairTable").querySelector("tbody");
    const rows = tableBody.rows;
    const data = [];

    for (let i = 0; i < rows.length; i++) {
      const date = rows[i].cells[1].querySelector("input").value;
      const repairSite = rows[i].cells[2].querySelector("input").value;
      const repairItem = rows[i].cells[3].querySelector("input").value;
      const quantity = rows[i].cells[4].querySelector("input").value;
      const unit = rows[i].cells[5].querySelector("input").value;
      const price = rows[i].cells[6].querySelector("input").value;

      data.push({ date, repairSite, repairItem, quantity, unit, price });
    }
    localStorage.setItem("repairData", JSON.stringify(data));
  }
  
  function loadTableData() {
    const storedData = localStorage.getItem("repairData");
    if (storedData) {
      const data = JSON.parse(storedData);
      // 清空现有表格内容（如果需要）
      const tableBody = document.getElementById("repairTable").querySelector("tbody");
      tableBody.innerHTML = "";
      // 根据缓存数据重建行
      data.forEach((rowData, index) => {
        createRow(index + 1, rowData.date);
        const currentRow = tableBody.rows[index];
        currentRow.cells[2].querySelector("input").value = rowData.repairSite;
        currentRow.cells[3].querySelector("input").value = rowData.repairItem;
        currentRow.cells[4].querySelector("input").value = rowData.quantity;
        currentRow.cells[5].querySelector("input").value = rowData.unit;
        currentRow.cells[6].querySelector("input").value = rowData.price;
        // 计算当前行的总价
        const quantity = parseFloat(rowData.quantity) || 0;
        const price = parseFloat(rowData.price) || 0;
        currentRow.cells[7].textContent = (quantity * price).toFixed(2);
      });
      updateSummary();
    }
  }
  
  loadTableData();
  
  const table = document
    .getElementById("repairTable")
    .getElementsByTagName("tbody")[0];
  const addRowButton = document.getElementById("addRow");
  const exportWordButton = document.getElementById("exportWord");
  const totalPriceCell = document.getElementById("totalPrice");

  function createRow(rowCount, previousDate = "") {
    const newRow = table.insertRow(-1);
    newRow.innerHTML = `
        <td data-label="序号">${rowCount}</td>
        <td data-label="日期"><input type="date" value="${previousDate}"></td>
        <td data-label="维修场所"><input type="text"></td>
        <td data-label="维修事项"><input type="text" required></td>
        <td data-label="数量"><input type="number" class="quantity" min="0" required></td>
        <td data-label="单位"><input type="text"></td>
        <td data-label="单价"><input type="number" class="price" min="0" required></td>
        <td class="total" data-label="总价">0.00</td>
        <td data-label="操作"><button class="delete-btn">删除</button></td>
    `;
    addEventListeners(newRow);
    updateSummary();
  }

  // 初始化1行
  createRow(1);

  // 添加新行
  addRowButton.addEventListener("click", function () {
    const rowCount = table.rows.length + 1;
    const lastRow = table.rows[table.rows.length - 1];
    const lastDateInput = lastRow.querySelector('input[type="date"]');
    const lastDateValue = lastDateInput ? lastDateInput.value : "";
    createRow(rowCount, lastDateValue);
    saveTableData();

    // 滚动到新添加的行
    setTimeout(() => {
      window.scrollTo(0, document.body.scrollHeight);
    }, 100);
  });

  function addEventListeners(row) {
    const quantityInput = row.querySelector(".quantity");
    const priceInput = row.querySelector(".price");
    const totalCell = row.querySelector(".total");
    const deleteBtn = row.querySelector(".delete-btn");

    quantityInput.addEventListener("input", updateTotal);
    priceInput.addEventListener("input", updateTotal);
    deleteBtn.addEventListener("click", deleteRow);

    function updateTotal() {
      const quantity = parseFloat(quantityInput.value) || 0;
      const price = parseFloat(priceInput.value) || 0;
      totalCell.textContent = (quantity * price).toFixed(2);
      updateSummary();
      saveTableData();
    }

    function deleteRow() {
      if (table.rows.length > 1) {
        row.remove();
        updateRowNumbers();
        updateSummary();
        saveTableData();
      } else {
        alert("至少保留一行数据");
      }
    }
  }

  function updateRowNumbers() {
    const rows = table.rows;
    for (let i = 0; i < rows.length; i++) {
      rows[i].cells[0].textContent = i + 1;
    }
  }

  function updateSummary() {
    let totalPrice = 0;
    const rows = table.rows;
    for (let i = 0; i < rows.length; i++) {
      const quantityInput = rows[i].querySelector(".quantity");
      const priceInput = rows[i].querySelector(".price");
      const quantity = parseFloat(quantityInput.value) || 0;
      const price = parseFloat(priceInput.value) || 0;
      totalPrice += quantity * price;
    }
    totalPriceCell.textContent = totalPrice.toFixed(2);
  }

  // 添加导出Word功能
  exportWordButton.addEventListener("click", exportToWord);

  // 添加验证函数
  function validateTable() {
    const rows = table.rows;
    let isValid = true;
    let errorMessage = "";

    for (let i = 0; i < rows.length; i++) {
      const repairItem = rows[i].querySelector('input[type="text"][required]');
      const quantity = rows[i].querySelector(".quantity");
      const price = rows[i].querySelector(".price");

      if (!repairItem || !repairItem.value.trim()) {
        errorMessage += `第 ${i + 1} 行的维修事项不能为\n`;
        isValid = false;
      }
      if (!quantity || !quantity.value.trim()) {
        errorMessage += `第 ${i + 1} 行的数量不能为空\n`;
        isValid = false;
      }
      if (!price || !price.value.trim()) {
        errorMessage += `第 ${i + 1} 行的单价不能为空\n`;
        isValid = false;
      }
    }

    return { isValid, errorMessage };
  }

  async function exportToWord() {
    const validation = validateTable();
    if (!validation.isValid) {
      alert("导出失败，请检查以下错误：\n" + validation.errorMessage);
      return;
    }

    const {
      Document,
      Paragraph,
      Table,
      TableRow,
      TableCell,
      HeadingLevel,
      AlignmentType,
      WidthType,
      BorderStyle,
      convertInchesToTwip,
    } = docx;

    // 创建文档
    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              margin: {
                top: convertInchesToTwip(1),
                right: convertInchesToTwip(1),
                bottom: convertInchesToTwip(1),
                left: convertInchesToTwip(1),
              },
            },
          },
          children: [
            new Paragraph({
              text: "空调维修明细",
              heading: HeadingLevel.HEADING_1,
              alignment: AlignmentType.CENTER,
              spacing: { before: 240, after: 240 },
              size: 36,
              color: "000000", // 设置标题颜色为黑色
            }),
            await createTable(),
          ],
        },
      ],
    });

    // 生成文档并保存
    docx.Packer.toBlob(doc).then((blob) => {
      saveAs(blob, "空调维修明细.docx");
    });
  }

  async function createTable() {
    const {
      Table,
      TableRow,
      TableCell,
      Paragraph,
      WidthType,
      BorderStyle,
      HeightRule,
    } = docx;

    const table = document.getElementById("repairTable");
    if (!table) {
      console.error("Table not found");
      return createErrorTable("Error: Table not found");
    }

    const rows = table.rows;
    if (!rows || rows.length === 0) {
      console.error("Table is empty");
      return createErrorTable("Error: Table is empty");
    }

    const tableRows = [];
    const columnData = [];

    // 添加表头
    const headerCells = Array.from(rows[0].cells)
      .slice(0, -1)
      .map((cell, index) => {
        columnData[index] = [];
        return {
          text: cell.innerText,
          cell: new TableCell({
            children: [new Paragraph({ text: cell.innerText, size: 28 })],
            shading: {
              fill: "CCCCCC",
            },
            verticalAlign: docx.VerticalAlign.CENTER,
          }),
        };
      });

    // 添加数据行
    for (let i = 1; i < rows.length - 1; i++) {
      const cells = Array.from(rows[i].cells)
        .slice(0, -1)
        .map((cell, index) => {
          const input = cell.querySelector("input");
          let value = input ? input.value : cell.innerText;

          // 如果当前单元格为空，尝试获取上面最近的非空值
          if (value.trim() === "") {
            for (let j = i - 1; j >= 1; j--) {
              const prevInput = rows[j].cells[index].querySelector("input");
              const prevValue = prevInput
                ? prevInput.value
                : rows[j].cells[index].innerText;
              if (prevValue.trim() !== "") {
                value = prevValue;
                break;
              }
            }
          }

          columnData[index].push(value);
          return new TableCell({
            children: [new Paragraph({ text: value, size: 26 })],
            verticalAlign: docx.VerticalAlign.CENTER,
          });
        });
      tableRows.push(
        new TableRow({
          children: cells,
          height: { value: 400, rule: HeightRule.ATLEAST },
        })
      );
    }

    // 移除空列
    const nonEmptyColumns = columnData.map((column, index) => {
      return column.some((cell) => cell.trim() !== "");
    });

    const filteredHeaderCells = headerCells
      .filter((_, index) => nonEmptyColumns[index])
      .map((item) => item.cell);

    const filteredTableRows = tableRows.map((row) => {
      if (row) {
        const filteredCells = row.root.filter(
          (cell, index) => nonEmptyColumns[index - 1]
        );
        return new TableRow({
          children: filteredCells,
          height: row.options.height,
        });
      }
      return row;
    });

    // 更新汇总行
    const nonEmptyColumnCount = nonEmptyColumns.filter(Boolean).length;
    const summaryRow = rows[rows.length - 1];
    const summaryCells = [
      new TableCell({
        children: [new Paragraph({ text: "汇总", size: 28 })],
        columnSpan: Math.max(1, nonEmptyColumnCount - 1),
        verticalAlign: docx.VerticalAlign.CENTER,
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: summaryRow.cells[summaryRow.cells.length - 2].innerText,
            size: 28,
          }),
        ],
        verticalAlign: docx.VerticalAlign.CENTER,
      }),
    ];

    const finalRows = [
      new TableRow({
        children: filteredHeaderCells,
        height: { value: 500, rule: HeightRule.ATLEAST },
      }),
      ...filteredTableRows,
      new TableRow({
        children: summaryCells,
        height: { value: 500, rule: HeightRule.ATLEAST },
      }),
    ];

    return new Table({
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
      rows: finalRows,
      borders: {
        top: { style: BorderStyle.SINGLE, size: 2 },
        bottom: { style: BorderStyle.SINGLE, size: 2 },
        left: { style: BorderStyle.SINGLE, size: 2 },
        right: { style: BorderStyle.SINGLE, size: 2 },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
        insideVertical: { style: BorderStyle.SINGLE, size: 1 },
      },
    });
  }

  function createErrorTable(errorMessage) {
    const { Table, TableRow, TableCell, Paragraph } = docx;
    return new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph({ text: errorMessage })],
            }),
          ],
        }),
      ],
    });
  }
});
