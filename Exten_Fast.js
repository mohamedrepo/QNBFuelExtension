(async function extractAllFuelTransactionsFast() {
    const FROM_DATE = '01/07/2023';
    const TO_DATE   = '30/06/2026';
    const DELAY = 500;

    const allPostedData  = [];
    const allPendingData = [];
    const allBalanceData = [];
    const emptyCards = [];
    const sleep = ms => new Promise(r => setTimeout(r, ms));

    const waitForModalOpen = async () => {
        const start = Date.now();
        while (Date.now() - start < 5000) {
            await sleep(200);
            const m = document.getElementById('myModal');
            if (m && m.style.display !== 'none' && (m.classList.contains('in') || m.getAttribute('aria-hidden') === 'false')) {
                await sleep(300);
                return true;
            }
        }
        return false;
    };

    const closeModalAndWait = async () => {
        const modal = document.getElementById('myModal');
        const btn = modal?.querySelector('.close, [data-dismiss="modal"]');
        if (btn) btn.click();
        else document.dispatchEvent(new KeyboardEvent('keydown', {key:'Escape'}));
        await sleep(800);
    };

    const setDate = async (fieldId, value) => {
        const input = document.getElementById(fieldId);
        if (!input) return;
        input.removeAttribute('readonly');
        input.removeAttribute('disabled');
        if (window.jQuery) {
            try { $(input).datepicker('setDate', value); } catch(e) {}
            $(input).val(value).trigger('change');
        }
        const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
        setter.call(input, value);
        input.dispatchEvent(new Event('input', {bubbles:true}));
        input.dispatchEvent(new Event('change', {bubbles:true}));
    };

    const clickSearch = () => {
        const btn = document.querySelector('#myModal input[type="submit"], #myModal button[type="submit"], #myModal .btn-search');
        if (btn) { btn.click(); return true; }
        return false;
    };

    const getOnlineBalance = async (cardRef) => {
        try {
            const btn = document.querySelector(`button[onclick*="RefreshCard(${cardRef}"]`);
            if (!btn) return '';
            btn.click();
            await sleep(2000);
            const msg = document.querySelector('#notificationModal .message')?.innerText;
            const match = msg?.match(/[\d,]+\.?\d*/);
            const bal = match ? match[0] : '';
            const closeBtn = document.querySelector('#notificationModal .close, #notificationModal .notification-btn-close');
            if (closeBtn) closeBtn.click();
            await sleep(300);
            return bal;
        } catch(e) { return ''; }
    };

    const extractAllRows = (tableId) => {
        try {
            if (window.jQuery && $.fn.dataTable?.isDataTable(`#${tableId}`)) {
                const rows = [];
                $(`#${tableId}`).DataTable().rows().every(function() {
                    const node = this.node();
                    if (node) {
                        const cols = node.querySelectorAll('td');
                        if (cols.length >= 4 && !cols[0].getAttribute('colspan')) {
                            rows.push([...cols].map(c => c.innerText?.trim()));
                        }
                    }
                });
                return rows;
            }
        } catch(e) {}
        const rows = [];
        document.getElementById(tableId)?.querySelectorAll('tbody tr').forEach(tr => {
            const cols = tr.querySelectorAll('td');
            if (cols.length >= 4 && !cols[0].getAttribute('colspan')) {
                rows.push([...cols].map(c => c.innerText?.trim()));
            }
        });
        return rows;
    };

    const waitForTable = async (tableId, timeoutMs = 6000) => {
        const start = Date.now();
        while (Date.now() - start < timeoutMs) {
            await sleep(300);
            try {
                if (window.jQuery && $.fn.dataTable?.isDataTable(`#${tableId}`)) {
                    const count = $(`#${tableId}`).DataTable().page.info().recordsTotal;
                    if (count > 0) return count;
                }
            } catch(e) {}
            const rows = document.getElementById(tableId)?.querySelectorAll('tbody tr') || [];
            if (rows.length > 0 && !rows[0].querySelector('td[colspan]')) {
                return rows.length;
            }
        }
        return 0;
    };

    console.log('='.repeat(60));
    console.log('FAST QNB FUEL EXTRACTOR');
    console.log(`Date Range: ${FROM_DATE} to ${TO_DATE}`);
    console.log('='.repeat(60));

    const cards = document.querySelectorAll('a.card-details');
    console.log(`Found ${cards.length} cards\n`);

    for (let i = 0; i < cards.length; i++) {
        const freshCards = document.querySelectorAll('a.card-details');
        if (i >= freshCards.length) break;

        const card = freshCards[i];
        const row = card.closest('tr');
        const cells = row.querySelectorAll('td');
        const cardNum = cells[0]?.innerText?.trim() || `Card_${i+1}`;
        const cardName = cells[1]?.innerText?.trim() || '';
        const tableBalance = cells[2]?.innerText?.trim() || '';
        const cardRef = (card.getAttribute('href')?.match(/cardRef=(\d+)/) || [])[1] || '';

        console.log(`[${i+1}/${cards.length}] ${cardNum}`);

        try {
            const onlineBalance = await getOnlineBalance(cardRef) || tableBalance;
            allBalanceData.push({card_number:cardNum, card_name:cardName, online_balance:onlineBalance, table_balance:tableBalance, card_ref:cardRef});

            card.click();
            await waitForModalOpen();

            await setDate('From', FROM_DATE);
            await sleep(50);
            await setDate('To', TO_DATE);
            await sleep(50);

            clickSearch();
            await waitForTable('dtPostedCardMovements');

            const pendingRows = extractAllRows('dtPindingCardMovements');
            const postedRows = extractAllRows('dtPostedCardMovements');

            pendingRows.forEach(cols => {
                allPendingData.push({
                    card_number:cardNum, card_name:cardName,
                    transaction_description:cols[0],
                    transaction_date:cols[1],
                    transaction_amount:cols[2]?.replace(/,/g,''),
                    debit_credit:cols[3]
                });
            });

            postedRows.forEach(cols => {
                allPostedData.push({
                    card_number:cardNum, card_name:cardName, online_balance:onlineBalance,
                    transaction_description:cols[0],
                    transaction_date:cols[1],
                    transaction_amount:cols[2]?.replace(/,/g,''),
                    transaction_type:cols[3]
                });
            });

            if (postedRows.length === 0) {
                emptyCards.push({card_number:cardNum, card_name:cardName, online_balance:onlineBalance});
            }

            console.log(`  ✓ Posted: ${postedRows.length}  Pending: ${pendingRows.length}`);

            await closeModalAndWait();
            await sleep(DELAY);

        } catch(err) {
            console.error(`  ✗ Error: ${err.message}`);
            try { await closeModalAndWait(); } catch(e) {}
        }
    }

    const sp = allPostedData.sort((a,b) => a.card_number.localeCompare(b.card_number)||a.transaction_type.localeCompare(b.transaction_type));
    const sn = allPendingData.sort((a,b) => a.card_number.localeCompare(b.card_number));
    const sb = allBalanceData.sort((a,b) => a.card_number.localeCompare(b.card_number));

    console.log(`\n✓ Posted: ${sp.length}  Pending: ${sn.length}  Balance: ${sb.length}  Empty: ${emptyCards.length}`);

    if (sp.length || sn.length || sb.length) {
        const cell  = (v,s='') => `<Cell${s?` ss:StyleID="${s}"`:''}><Data ss:Type="String">${String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</Data></Cell>`;
        const hcell = v => cell(v,'h');
        const erow  = (cells,i) => `<Row${i%2?' ss:StyleID="e"':''}>${cells}</Row>`;

        let html = '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>';
        html += '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">';
        html += '<Styles><Style ss:ID="h"><Font ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#1d2552" ss:Pattern="Solid"/></Style><Style ss:ID="b"><Font ss:Bold="1" ss:Color="#1d2552"/></Style><Style ss:ID="e"><Interior ss:Color="#f2f2f2" ss:Pattern="Solid"/></Style></Styles>';

        html += '<Worksheet ss:Name="Posted Transactions"><Table>';
        html += '<Row>'+['Card Number','Card Name','Online Balance','Transaction Date','Transaction Description','Amount','Type (D/C)'].map(hcell).join('')+'</Row>';
        sp.forEach((d,i) => html += erow(cell(d.card_number)+cell(d.card_name)+cell(d.online_balance,'b')+cell(d.transaction_date)+cell(d.transaction_description)+cell(d.transaction_amount)+cell(d.transaction_type),i));
        html += '</Table></Worksheet>';

        html += '<Worksheet ss:Name="Pending Transactions"><Table>';
        html += '<Row>'+['Card Number','Card Name','Transaction Date','Transaction Description','Amount','Debit/Credit'].map(hcell).join('')+'</Row>';
        sn.forEach((d,i) => html += erow(cell(d.card_number)+cell(d.card_name)+cell(d.transaction_date)+cell(d.transaction_description)+cell(d.transaction_amount)+cell(d.debit_credit),i));
        html += '</Table></Worksheet>';

        html += '<Worksheet ss:Name="Online Balance"><Table>';
        html += '<Row>'+['Card Number','Card Name','Online Balance','Table Balance','Card Ref'].map(hcell).join('')+'</Row>';
        sb.forEach((d,i) => html += erow(cell(d.card_number)+cell(d.card_name)+cell(d.online_balance,'b')+cell(d.table_balance)+cell(d.card_ref),i));
        html += '</Table></Worksheet>';

        html += '<Worksheet ss:Name="Empty Cards"><Table>';
        html += '<Row>'+['Card Number','Card Name','Online Balance'].map(hcell).join('')+'</Row>';
        emptyCards.forEach((d,i) => html += erow(cell(d.card_number)+cell(d.card_name)+cell(d.online_balance,'b'),i));
        html += '</Table></Worksheet>';

        html += '</Workbook>';

        const blob = new Blob([html], {type:'application/vnd.ms-excel'});
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `QNB_Fuel_${FROM_DATE.replace(/\//g,'')}_to_${TO_DATE.replace(/\//g,'')}.xls`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        console.log(`\n✅ File: ${link.download}`);
        console.log(`📊 Posted: ${sp.length}  ⏳ Pending: ${sn.length}  💰 Balance: ${sb.length}  ❌ Empty: ${emptyCards.length}`);
    } else {
        console.log('❌ No data extracted');
    }
    console.log('='.repeat(60));
})();