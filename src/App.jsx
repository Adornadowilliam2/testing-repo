import React, { useState } from 'react';
import * as XLSX from 'xlsx';

export default function App() {
  const [sqlOutput, setSqlOutput] = useState('');
  const [selectedColor, setSelectedColor] = useState('16113331');

  const STARTING_DEPT_ID = 1;  
  const STARTING_CAT_ID = 1;   

  const colorOptions = [
    { label: "Snow (Default)", value: "16113331" },
    { label: "Brown", value: "4279205" },
    { label: "Light Red", value: "8421631" },
    { label: "Light Green", value: "9498256" },
    { label: "Green", value: "3931654" },
    { label: "Light Blue", value: "14596231" },
    { label: "Royal Blue", value: "16747520" },
    { label: "Coffee", value: "1993170" },
    { label: "Firewall", value: "2237106" },
    { label: "Salmon", value: "7490810" },
    { label: "Crimson", value: "3936508" },
    { label: "Tomato", value: "4679935" },
    { label: "Non-Special", value: "10025880" },
    { label: "Special", value: "4178381" },
    { label: "Lavender", value: "16443134" },
    { label: "Light Orange", value: "8179967" },
    { label: "Mint", value: "13214781" },
    { label: "Honeydew", value: "15794160" },
    { label: "Light Pink", value: "12695295" },
    { label: "Powder Blue", value: "15130800" },
    { label: "Linen", value: "15132410" },
    { label: "Plum", value: "14524637" },
    { label: "Cyan", value: "16776960" },
    { label: "Turquoise", value: "13688896" },
    { label: "Yellow", value: "14745599" },
    { label: "Lemon Chiffon", value: "13499135" },
    { label: "Aquamarine", value: "13959039" },
    { label: "Chemical", value: "8427616" }
  ];

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);
      
      processExcelToSql(jsonData);
    };
    reader.readAsArrayBuffer(file);
  };

  const processExcelToSql = (rows) => {
    if (rows.length === 0) {
      setSqlOutput('-- No rows discovered in the uploaded sheet file.');
      return;
    }

    const structuralDeptMap = new Map(); 
    const structuralCatMap = new Map();  

    let structDeptId = STARTING_DEPT_ID;
    let structCatId = STARTING_CAT_ID;

    const codeButtonCounters = new Map();
    const itemButtonCounters = new Map();
    const activeBackColor = parseInt(selectedColor, 10);

    // FIRST PASS: ISOLATE DEPARTMENTS
    rows.forEach((row) => {
      const deptText = (row.dept || row['dept'] || row['Department'] || '').toString().trim();
      if (deptText && !structuralDeptMap.has(deptText)) {
        const paddedCode = String(structDeptId).padStart(4, '0');
        structuralDeptMap.set(deptText, { id: structDeptId, code: paddedCode });
        codeButtonCounters.set(paddedCode, 1);
        structDeptId++;
      }
    });

    // SECOND PASS: MAP UNIQUE CATEGORIES
    rows.forEach((row) => {
      const deptText = (row.dept || row['dept'] || row['Department'] || '').toString().trim();
      const catText = (row.category || row['category'] || row['Category'] || '').toString().trim();

      if (!deptText || !catText) return;

      const deptInfo = structuralDeptMap.get(deptText);
      const comboKey = `${deptText}|${catText}`;

      if (deptInfo && !structuralCatMap.has(comboKey)) {
        const targetCode = deptInfo.code;
        const currentButtonIndex = codeButtonCounters.get(targetCode);
        const paddedCatCode = String(structCatId).padStart(4, '0');

        structuralCatMap.set(comboKey, { 
          id: structCatId, 
          code: paddedCatCode,
          name: catText, 
          deptName: deptText,
          buttonIndex: currentButtonIndex 
        });

        itemButtonCounters.set(structCatId, 1);
        codeButtonCounters.set(targetCode, currentButtonIndex + 1);
        structCatId++;
      }
    });

    // THIRD PASS: GENERATE SCHEMATIC SCRIPT
    let script = `-- ===================================================\n`;
    script += `-- DYNAMIC POS MASTERFILE CONVERSION SCRIPT\n`;
    script += `-- Target Database: [susfood]\n`;
    script += `-- Generated on: 2026-05-19\n`;
    script += `-- Selected Layout BackColor Code: ${activeBackColor}\n`;
    script += `-- ===================================================\n\n`;

    // 1. Department Inserts
    if (structuralDeptMap.size > 0) {
      script += `-- 1. DEPARTMENTS\n`;
      script += `SET IDENTITY_INSERT [susfood].[dbo].[dept] ON;\n`;
      structuralDeptMap.forEach((info, name) => {
        script += `INSERT INTO [susfood].[dbo].[dept] (dept, deptd, button, backcolor, forecolor) VALUES (${info.id}, '${name.replace(/'/g, "''")}', 'menu${info.id}', ${activeBackColor}, 0);\n`;
      });
      script += `SET IDENTITY_INSERT [susfood].[dbo].[dept] OFF;\n\n`;

      script += `SET IDENTITY_INSERT [susfood].[dbo].[dept2] ON;\n`;
      structuralDeptMap.forEach((info, name) => {
        script += `INSERT INTO [susfood].[dbo].[dept2] (dept, deptd, button, backcolor, forecolor, dtrandate, button_i) VALUES ('${info.code}', '${name.replace(/'/g, "''")}', 'menu${info.id}', ${activeBackColor}, 0, '1900-01-01 00:00:00.000', 0);\n`;
      });
      script += `SET IDENTITY_INSERT [susfood].[dbo].[dept2] OFF;\n\n`;
    }

    // 2. Category Inserts
    if (structuralCatMap.size > 0) {
      script += `-- 2. CATEGORIES\n`;
      script += `SET IDENTITY_INSERT [susfood].[dbo].[category] ON;\n`;
      structuralCatMap.forEach((catInfo, key) => {
        const deptInfo = structuralDeptMap.get(catInfo.deptName);
        if (deptInfo) {
          script += `INSERT INTO [susfood].[dbo].[category] (category, categoryd, dept, button, backcolor, forecolor, min, max, isbulk) VALUES (${catInfo.id}, '${catInfo.name.replace(/'/g, "''")}', '${deptInfo.code}', 'submenu${catInfo.buttonIndex}', ${activeBackColor}, 0, 0.00, 0.00, 0);\n`;
        }
      });
      script += `SET IDENTITY_INSERT [susfood].[dbo].[category] OFF;\n\n`;

      script += `SET IDENTITY_INSERT [susfood].[dbo].[category2] ON;\n`;
      structuralCatMap.forEach((catInfo, key) => {
        const deptInfo = structuralDeptMap.get(catInfo.deptName);
        if (deptInfo) {
          script += `INSERT INTO [susfood].[dbo].[category2] (category, categoryd, dept, button, backcolor, forecolor, dtrandate, button_i, addon) VALUES (${catInfo.id}, '${catInfo.name.replace(/'/g, "''")}', '${deptInfo.code}', 'submenu${catInfo.buttonIndex}', ${activeBackColor}, 0, '1900-01-01 00:00:00.000', 0, 0);\n`;
        }
      });
      script += `SET IDENTITY_INSERT [susfood].[dbo].[category2] OFF;\n\n`;
    }

    // 3. Items Processing
    script += `-- 3. ITEMS DATA MAP\n`;
    rows.forEach((row, index) => {
      const itemCode = (row.itemcode || row['itemcode'] || '').toString().trim();
      const barcode = (row.barcode || row['barcode'] || itemCode).toString().trim();
      const desc = (row.desc || row['desc'] || '').toString().replace(/'/g, "''").trim();
      const posdesc = (row['post desc'] || row['posdesc'] || desc).toString().replace(/'/g, "''").trim();
      
      const rawPrice = (row.sell || row['selling price'] || '0').toString().replace(/[^0-9.]/g, '');
      const parsedPrice = parseFloat(rawPrice);
      const price = isNaN(parsedPrice) ? 0.00 : Number(parsedPrice.toFixed(2));
      
      const statusValue = (row.status || row['item status'] || '').toString().toUpperCase().trim() === 'ACTIVE' ? 1 : 0;

      const deptText = (row.dept || row['dept'] || '').toString().trim();
      const catText = (row.category || row['category'] || '').toString().trim();

      const deptInfo = structuralDeptMap.get(deptText);
      const comboKey = `${deptText}|${catText}`;
      const catInfo = structuralCatMap.get(comboKey);

      if (deptInfo && catInfo) {
        const currentItemButtonNum = itemButtonCounters.get(catInfo.id);
        const cbButtonString = `cb${currentItemButtonNum}`;
        itemButtonCounters.set(catInfo.id, currentItemButtonNum + 1);

        script += `/* Excel Row ${index + 2} - Item: ${itemCode} */\n`;

        // ===================================================
        // 1. [items] SCHEMA MAP 
        // ===================================================
        const itemsSchemaMap = {
          "prodid": 0, "itemcode": `'${itemCode}'`, "barcode": `'${barcode}'`, "[desc]": `'${desc}'`, "posdesc": `'${posdesc}'`,
          "dept": deptInfo.id, "category": catInfo.id, "subcate": 1, "brand": 1, "area": 1,
          "color": 1, "size": 1, "costp": 0.00000, "costs": 0.00, "costi": 0.00,
          "costd": 0.00000, "costp1": 0.00, "sell": price, "markup": 0.0000, "min": 0.00,
          "max": 0.00, "mdiscount": "''", "discamt": 0.00, "sdate": "'1900-01-01 00:00:00.000'", "stime": "''",
          "edate": "'1900-01-01 00:00:00.000'", "etime": "''", "casepack": 0.00, "suppcode": 0, "opend": 0,
          "datearive": "GETDATE()", "currency": 1, "remarks": "''", "expflag": 0, "qty1": 0,
          "qty2": 0, "qty3": 0, "qty4": 0, "qty5": 0, "levelp1": 0.00,
          "levelp2": 0.00, "levelp3": 0.00, "levelp4": 0.00, "levelp5": 0.00, "olditem": 0,
          "deffect": "GETDATE()", "bartype": 1, "status": statusValue, "picture": "''", "compose": 0,
          "dcost": "''", "dsell": "''", "lnonsale": 0, "vat": 0.00, "addoncost": 0.00,
          "cubic": 0.000, "weight": 0.000, "active": 1, "lsurcharge": 0, "surchper": 0.00,
          "discl1": 0.00, "discl2": 0.00, "discl3": 0.00, "discl4": 0.00, "discl5": 0.00,
          "button": `'${cbButtonString}'`, "itemtype": 1, "combo": 0, "backcolor": activeBackColor, "forecolor": 0,
          "rawitem": 0, "computed": 0, "bracket1": "''", "bracket2": "''", "bracket3": "''",
          "bracket4": "''", "bracket5": "''", "lackserial": 0, "[group]": 1, "divsn": 1,
          "stocktran": 0, "sell1": 0.00, "sell2": 0.00, "sell3": 0.00, "unit": 1,
          "itemstat": 1, "wunit": 1, "picbutton": 0, "lnonvat": 0, "avgcost": 0.00,
          "shelflife": 0, "printer": "''", "allloc": 0, "pickfrom": 0, "bom": 0,
          "floorstock": 0, "pmcode": 0, "ordmin": 0, "ordmax": 0, "yieldpct": 0.00,
          "avgdate": "'1900-01-01 00:00:00.000'", "stddate": "'1900-01-01 00:00:00.000'", "catbutton": "''", "sellunit": 0, "convqty": 0.0000,
          "promocode": "''", "lweighitem": 0, "lbatchno": 0, "tareweight": 0.00, "sell4": 0.00,
          "laskserial": 0, "proctime": 0.00, "kdsdelay": 0.00, "nosurch": 0, "refundable": 0,
          "linactive": 0, "laskmod": 0, "lcutline": 0, "perishable": 0, "indented": 0,
          "lnolt": 0, "pcountprior": 0, "dailypcount": 0, "wastage": 0.0000, "stockalert": 0,
          "lweight": 0, "lcake": 0, "dept2": `'${deptInfo.code}'`, "category2": `'${catInfo.code}'`,
          "printer1": "''", "printer2": "''", "printer3": "''", "printer4": "''", "fenixcode": "''",
          "tkeoutcode": "''", "qbcode": "''", "calories": "''", "nutrifacts": "''", "cboremarks": "''",
          "lsample": 0
        };

        // ===================================================
        // 2. [plupos] SCHEMA MAP
        // ===================================================
        const pluposSchemaMap = {
          "prodid": 0, "itemcode": `'${itemCode}'`, "barcode": `'${barcode}'`, "[desc]": `'${desc}'`, "posdesc": `'${posdesc}'`,
          "dept": deptInfo.id, "category": catInfo.id, "subcate": 1, "brand": 1, "area": 1,
          "color": 1, "size": 1, "costp": 0.00000, "costd": 0.00000, "costp1": 0.00,
          "markup": 0.0000, "min": 0.00, "max": 0.00, "discamt": 0.00, "sdate": "'1900-01-01 00:00:00.000'",
          "stime": "''", "edate": "'1900-01-01 00:00:00.000'", "etime": "''", "suppcode": 0, "opend": 0,
          "currency": 1, "remarks": "''", "levelp1": 0.00, "levelp2": 0.00, "levelp3": 0.00,
          "levelp4": 0.00, "levelp5": 0.00, "deffect": "GETDATE()", "bartype": 1, "status": statusValue,
          "picture": "''", "compose": 0, "lnonsale": 0, "vat": 0.00, "addoncost": 0.00,
          "cubic": 0.000, "weight": 0.000, "active": 1, "lsurcharge": 0, "surchper": 0.00,
          "discl1": 0.00, "discl2": 0.00, "discl3": 0.00, "discl4": 0.00, "discl5": 0.00,
          "button": `'${cbButtonString}'`, "itemtype": 1, "combo": 0, "backcolor": activeBackColor, "forecolor": 0,
          "rawitem": 0, "computed": 0, "bracket1": "''", "bracket2": "''", "bracket3": "''",
          "bracket4": "''", "bracket5": "''", "lackserial": 0, "[group]": 1, "divsn": 1,
          "stocktran": 0, "unit": 1, "itemstat": 1, "wunit": 1, "lnonvat": 0,
          "avgcost": 0.00, "shelflife": 0, "printer": "''", "allloc": 0, "pickfrom": 0,
          "bom": 0, "catbutton": "''", "sellunit": 0, "convqty": 0.0000, "promocode": "''",
          "lweighitem": 0, "lbatchno": 0, "tareweight": 0.00, "laskserial": 0, "proctime": 0.00,
          "kdsdelay": 0.00, "nosurch": 0, "refundable": 0, "linactive": 0, "laskmod": 0,
          "lcutline": 0, "perishable": 0, "indented": 0, "lnolt": 0, "pcountprior": 0,
          "dailypcount": 0, "wastage": 0.0000, "stockalert": 0, "lcake": 0, "dept2": `'${deptInfo.code}'`,
          "category2": `'${catInfo.code}'`, "printer1": "''", "printer2": "''", "printer3": "''", "fenixcode": "''",
          "tkeoutcode": "''", "qbcode": "''", "calories": "''", "nutrifacts": "''", "cboremarks": "''",
          "lsample": 0
        };

        // ===================================================
        // 3. [bitems] SCHEMA MAP
        // ===================================================
        const bitemsSchemaMap = {
          "prodid": 0, "itemcode": `'${itemCode}'`, "barcode": `'${barcode}'`, "[desc]": `'${desc}'`, "posdesc": `'${posdesc}'`,
          "dept": deptInfo.id, "category": catInfo.id, "subcate": 1, "brand": 1, "area": 1,
          "color": 1, "size": 1, "costp": 0.00000, "costs": 0.00, "costi": 0.00,
          "costd": 0.00000, "costp1": 0.00, "sell": price, "markup": 0.0000, "min": 0.00,
          "max": 0.00, "mdiscount": "''", "discamt": 0.00, "sdate": "'1900-01-01 00:00:00.000'",
          "stime": "''", "edate": "'1900-01-01 00:00:00.000'", "etime": "''", "casepack": 0.00, "suppcode": 1,
          "opend": 0, "datearive": "GETDATE()", "currency": 1, "remarks": "''", "expflag": 0,
          "qty1": 0, "qty2": 0, "qty3": 0, "qty4": 0, "qty5": 0,
          "levelp1": 0.00, "levelp2": 0.00, "levelp3": 0.00, "levelp4": 0.00, "levelp5": 0.00,
          "olditem": 0, "deffect": "GETDATE()", "bartype": 1, "status": statusValue, "picture": "''",
          "compose": 0, "dcost": "''", "dsell": "''", "lnonsale": 0, "vat": 0.00,
          "addoncost": 0.00, "cubic": 0.000, "weight": 0.000, "active": 1, "lsurcharge": 0,
          "surchper": 0.00, "discl1": 0.00, "discl2": 0.00, "discl3": 0.00, "discl4": 0.00,
          "discl5": 0.00, "button": `'${cbButtonString}'`, "itemtype": 1, "printer": "''", "combo": 0,
          "backcolor": activeBackColor, "forecolor": 0, "rawitem": 0, "computed": 0, "bracket1": "''",
          "bracket2": "''", "bracket3": "''", "bracket4": "''", "bracket5": "''", "lackserial": 0,
          "lnonvat": 0, "[group]": 1, "divsn": 1, "stocktran": 0, "sell1": 0.00,
          "sell2": 0.00, "sell3": 0.00, "unit": 1, "location": 2, "avgcost": 0.00,
          "shelflife": 0, "pickfrom": 0, "bom": 0, "catbutton": "''", "sellunit": 1,
          "convqty": 1.0000, "lweighitem": 0, "lbatchno": 0, "tareweight": 0.00, "sell4": 0.00,
          "printer1": "''", "printer2": "''", "printer3": "''", "fenixcode": "''", "tkeoutcode": "''",
          "sapcode": "''"
        };

        // ===================================================
        // 4. [lastean8] SCHEMA MAP
        // ===================================================
        const lastean8SchemaMap = {
          "barcode": `'${barcode}'`
        };

        // ===================================================
        // 5. [addeditlog] SCHEMA MAP (Hardcoded to Chase)
        // ===================================================
        const addeditlogSchemaMap = {
          "barcode": `'${barcode}'`,
          "[desc]": `'${desc}'`,
          "dept": `'${deptInfo.code}'`, 
          "category": `'${catInfo.code}'`,
          "remarks": `'ADD : ${desc}'`,
          "date": "CONVERT(varchar(10), GETDATE(), 120) + ' 00:00:00.000'",
          "time": "CONVERT(varchar(8), GETDATE(), 108)",
          "[user]": "'Chase'"
        };

        // Automatic generation arrays compilation
        const itemsColumns = Object.keys(itemsSchemaMap).join(', ');
        const itemsValues = Object.values(itemsSchemaMap).join(', ');

        const pluposColumns = Object.keys(pluposSchemaMap).join(', ');
        const pluposValues = Object.values(pluposSchemaMap).join(', ');

        const bitemsColumns = Object.keys(bitemsSchemaMap).join(', ');
        const bitemsValues = Object.values(bitemsSchemaMap).join(', ');

        const lastean8Columns = Object.keys(lastean8SchemaMap).join(', ');
        const lastean8Values = Object.values(lastean8SchemaMap).join(', ');

        const addeditlogColumns = Object.keys(addeditlogSchemaMap).join(', ');
        const addeditlogValues = Object.values(addeditlogSchemaMap).join(', ');

        // Push clean SQL mutations into result collector
        script += `INSERT INTO [susfood].[dbo].[items] (${itemsColumns}) VALUES (${itemsValues});\n`;
        script += `INSERT INTO [susfood].[dbo].[plupos] (${pluposColumns}) VALUES (${pluposValues});\n`;
        script += `INSERT INTO [susfood].[dbo].[bitems] (${bitemsColumns}) VALUES (${bitemsValues});\n`;
        script += `INSERT INTO [susfood].[dbo].[lastean8] (${lastean8Columns}) VALUES (${lastean8Values});\n`;
        script += `INSERT INTO [susfood].[dbo].[addeditlog] (${addeditlogColumns}) VALUES (${addeditlogValues});\n\n`;
      }
    });

    setSqlOutput(script);
  };

  return (
    <div style={{ padding: '24px', fontFamily: 'sans-serif', maxWidth: '1000px', margin: '0 auto' }}>
      <h2 style={{ marginBottom: '4px' }}>POS Masterfile Script Generator ([susfood])</h2>
      <p style={{ color: '#666', marginTop: '0' }}>Select your layout skin configuration context and drop the raw source template file.</p>
       
      <div style={{ margin: '20px 0 10px 0', display: 'flex', flexDirection: 'column', gap: '6px' }}>
        <label style={{ fontSize: '14px', fontWeight: 'bold', color: '#444' }}>Select Grid Layout Button Color:</label>
        <select 
          value={selectedColor} 
          onChange={(e) => setSelectedColor(e.target.value)}
          style={{ padding: '10px', fontSize: '14px', borderRadius: '4px', border: '1px solid #ccc', maxWidth: '300px', cursor: 'pointer', background: '#fff' }}
        >
          {colorOptions.map((opt) => (
            <option key={opt.value} value={opt.value}>{opt.label} ({opt.value})</option>
          ))}
        </select>
      </div>

      <div style={{ padding: '20px', background: '#fcfcfc', border: '2px dashed #bbb', borderRadius: '8px', textAlign: 'center', margin: '15px 0' }}>
        <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} style={{ fontSize: '15px' }} />
      </div>

      {sqlOutput && (
        <div style={{ marginTop: '24px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
            <h3 style={{ margin: 0 }}>SSMS Ready Query Script</h3>
            <button 
              onClick={() => navigator.clipboard.writeText(sqlOutput)}
              style={{ padding: '8px 16px', background: '#007acc', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}
            >
              Copy Script
            </button>
          </div>
          <textarea
            value={sqlOutput}
            readOnly
            style={{ width: '100%', height: '450px', fontFamily: 'Consolas, Monaco, monospace', fontSize: '13px', padding: '14px', backgroundColor: '#1e1e1e', color: '#d4d4d4', border: '1px solid #333', borderRadius: '6px', overflowX: 'auto', whiteSpace: 'pre' }}
          />
        </div>
      )}
    </div>
  );
}