(async function extractAllFuelTransactionsThreeSheets() {
    const FROM_DATE = '01/07/2023';
    const TO_DATE   = '30/06/2026';
    const DELAY = 2000;

    const allPostedData  = [];
    const allPendingData = [];
    const allBalanceData = [];
    const emptyCards = [];
    const sleep = ms => new Promise(r => setTimeout(r, ms));

    // ── XHR intercept + promise tracking ──────────────────────
    const _open = XMLHttpRequest.prototype.open;
    const _send = XMLHttpRequest.prototype.send;
    let pendingXHR = null;

    XMLHttpRequest.prototype.open = function(method, url, ...args) {
        this._url = url; this._method = method;
        return _open.apply(this, [method, url, ...args]);
    };
    XMLHttpRequest.prototype.send = function(body) {
        if (this._method === 'POST' && this._url?.includes('/FuelCard/Home/Details') && typeof body === 'string') {
            const p = new URLSearchParams(body);
            p.set('From', FROM_DATE);
            p.set('To',   TO_DATE);
            body = p.toString();
            console.log(`  🔄 Intercepted → From:${FROM_DATE} To:${TO_DATE}`);

            // Create a new promise for THIS specific XHR
            const xhrPromise = new Promise(resolve => {
                this.addEventListener('load',  () => { console.log(`  ✅ XHR done (${this.status})`); resolve(this.responseText); });
                this.addEventListener('error', () => { console.log(`  ✗ XHR error`); resolve(null); });
                this.addEventListener('abort', () => { console.log(`  ✗ XHR aborted`); resolve(null); });
            });
            pendingXHR = xhrPromise;
        }
        return _send.apply(this, [body]);
    };

    // Wait for the CURRENT pending XHR then extra settle time
    const waitForCurrentXHR = async (settleMs = 2000) => {
        if (pendingXHR) {
            await pendingXHR;
            pendingXHR = null;
        }
        await sleep(settleMs); // let DataTables process the response
    };

    // Wait for table to stabilize
    const waitForTableStable = async (tableId, timeoutMs = 10000) => {
        const start = Date.now();
        let lastCount = -1;
        let stableFor = 0;
        while (Date.now() - start < timeoutMs) {
            await sleep(400);
            let count = 0;
            try {
                if (window.jQuery && $.fn.dataTable?.isDataTable(`#${tableId}`)) {
                    count = $(`#${tableId}`).DataTable().page.info().recordsTotal;
                } else {
                    const rows = document.getElementById(tableId)?.querySelectorAll('tbody tr') || [];
                    count = (rows.length === 1 && rows[0].querySelector('td[colspan]')) ? 0 : rows.length;
                }
            } catch(e) {}
            if (count === lastCount) {
                stableFor += 400;
                if (count > 0  && stableFor >= 1600) { console.log(`  ✓ #${tableId}: ${count} records`); return count; }
                if (count === 0 && stableFor >= 3200) { console.log(`  ✓ #${tableId}: empty`); return 0; }
            } else {
                stableFor = 0;
                lastCount = count;
            }
        }
        return lastCount;
    };

    const showAllRows = async (tableId) => {
        try {
            if (window.jQuery && $.fn.dataTable?.isDataTable(`#${tableId}`)) {
                $(`#${tableId}`).DataTable().page.len(-1).draw(false);
                await sleep(1000);
                return;
            }
        } catch(e) {}
        const dd = document.querySelector(`select[name="${tableId}_length"]`);
        if (dd) { dd.value = '-1'; dd.dispatchEvent(new Event('change', {bubbles:true})); await sleep(1000); }
    };

    const extractAllRows = (tableId) => {
        const rows = [];
        try {
            if (window.jQuery && $.fn.dataTable?.isDataTable(`#${tableId}`)) {
                $(`#${tableId}`).DataTable().rows().every(function() {
                    const node = this.node();
                    if (node) {
                        const cols = node.querySelectorAll('td');
                        if (cols.length >= 4 && !cols[0].getAttribute('colspan')) {
                            rows.push([...cols].map(c => c.innerText?.trim()));
                        }
                    }
                });
                if (rows.length) return rows;
            }
        } catch(e) {}
        document.getElementById(tableId)?.querySelectorAll('tbody tr').forEach(tr => {
            const cols = tr.querySelectorAll('td');
            if (cols.length >= 4 && !cols[0].getAttribute('colspan')) {
                rows.push([...cols].map(c => c.innerText?.trim()));
            }
        });
        return rows;
    };

    // Close modal and WAIT until it's fully gone from DOM
    const closeModalAndWait = async () => {
        const modal = document.getElementById('myModal');
        const btn = modal?.querySelector('.close, [data-dismiss="modal"]');
        if (btn) btn.click();
        else document.dispatchEvent(new KeyboardEvent('keydown', {key:'Escape'}));

        // Wait until modal is hidden (Bootstrap removes 'in' class and display becomes none)
        const start = Date.now();
        while (Date.now() - start < 5000) {
            await sleep(200);
            const m = document.getElementById('myModal');
            const isHidden = !m ||
                             m.style.display === 'none' ||
                             !m.classList.contains('in') ||
                             m.getAttribute('aria-hidden') === 'true';
            if (isHidden) {
                console.log('  ✓ Modal closed');
                await sleep(500); // extra buffer after modal gone
                return;
            }
        }
        console.log('  ⚠ Modal close timeout');
        await sleep(500);
    };

    // Wait for modal to be FULLY open
    const waitForModalOpen = async () => {
        const start = Date.now();
        while (Date.now() - start < 8000) {
            await sleep(300);
            const m = document.getElementById('myModal');
            const isOpen = m &&
                           m.style.display !== 'none' &&
                           (m.classList.contains('in') || m.getAttribute('aria-hidden') === 'false');
            if (isOpen) {
                await sleep(500);
                return true;
            }
        }
        console.log('  ⚠ Modal open timeout');
        return false;
    };

    const closeNotificationModal = async () => {
        const btn = document.querySelector('#notificationModal .close, #notificationModal .notification-btn-close');
        if (btn) { btn.click(); await sleep(500); }
    };

    const setDate = async (fieldId, value) => {
        const input = document.getElementById(fieldId);
        if (!input) return;
        input.removeAttribute('readonly');
        input.removeAttribute('disabled');
        if (window.jQuery) {
            try {
                $(input).datepicker('option', 'minDate', new Date(2000,0,1));
                $(input).datepicker('option', 'maxDate', new Date(2030,11,31));
                $(input).datepicker('setDate', value);
            } catch(e) {}
            $(input).val(value).trigger('change');
        }
        const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
        setter.call(input, value);
        input.dispatchEvent(new Event('input',  {bubbles:true}));
        input.dispatchEvent(new Event('change', {bubbles:true}));
    };

    const clickSearch = () => {
        const btn = document.querySelector(
            '#myModal input[type="submit"], #myModal button[type="submit"], ' +
            '#myModal .btn-search, .modal-content input[type="submit"]'
        );
        if (btn) { btn.click(); return true; }
        return false;
    };

    const getOnlineBalance = async (cardRef) => {
        try {
            const btn = document.querySelector(`button[onclick*="RefreshCard(${cardRef}"]`);
            if (!btn) return null;
            btn.click();
            await sleep(2500);
            const msg   = document.querySelector('#notificationModal .message')?.innerText;
            const match = msg?.match(/[\d,]+\.?\d*/);
            const bal   = match ? match[0] : '';
            await closeNotificationModal();
            await sleep(500);
            return bal;
        } catch(e) { return null; }
    };

    const parseDateForSort = s => {
        if (!s) return '00000000';
        const p = s.match(/(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})/);
        if (p) { let[_,d,m,y]=p; if(y.length===2)y='20'+y; return `${y}${m.padStart(2,'0')}${d.padStart(2,'0')}`; }
        return s;
    };
    const sortPosted  = d => d.sort((a,b) => a.card_number.localeCompare(b.card_number)||a.transaction_type.localeCompare(b.transaction_type)||parseDateForSort(a.transaction_date).localeCompare(parseDateForSort(b.transaction_date)));
    const sortPending = d => d.sort((a,b) => a.card_number.localeCompare(b.card_number)||parseDateForSort(a.transaction_date).localeCompare(parseDateForSort(b.transaction_date)));
    const sortBalance = d => d.sort((a,b) => a.card_number.localeCompare(b.card_number));

    console.log('='.repeat(70));
    console.log('QNB FUEL CARD EXTRACTOR — MODAL AWAIT VERSION');
    console.log(`Date Range: ${FROM_DATE} to ${TO_DATE}`);
    console.log('='.repeat(70));

    const cards = document.querySelectorAll('a.card-details');
    console.log(`Found ${cards.length} cards\n`);

    for (let i = 0; i < cards.length; i++) {
        try {
            const freshCards = document.querySelectorAll('a.card-details');
            if (i >= freshCards.length) break;

            const card   = freshCards[i];
            const row    = card.closest('tr');
            const cells  = row.querySelectorAll('td');
            const cardNum      = cells[0]?.innerText?.trim() || `Card_${i+1}`;
            const cardName     = cells[1]?.innerText?.trim() || '';
            const tableBalance = cells[2]?.innerText?.trim() || '';
            const cardRef      = (card.getAttribute('href')?.match(/cardRef=(\d+)/) || [])[1] || '';

            console.log(`\n[${i+1}/${cards.length}] ${cardNum}`);

            const onlineBalance = await getOnlineBalance(cardRef) || tableBalance;
            allBalanceData.push({card_number:cardNum, card_name:cardName, online_balance:onlineBalance, table_balance:tableBalance, card_ref:cardRef});

            // Reset pending XHR tracker before opening modal
            pendingXHR = null;

            // Open modal and wait until fully visible
            card.click();
            await waitForModalOpen();

            // Set dates
            await setDate('From', FROM_DATE);
            await sleep(200);
            await setDate('To', TO_DATE);
            await sleep(200);

            // Click search — XHR promise is created inside the interceptor
            clickSearch();

            // Wait for XHR to fully complete + DataTables to process
            console.log(`  ⏳ Waiting for XHR...`);
            await waitForCurrentXHR(2000);

            // Now wait for table to stabilize
            await waitForTableStable('dtPostedCardMovements');

            // Expand to show all
            await showAllRows('dtPostedCardMovements');
            await sleep(500);

            // Extract PENDING
            let pendingCount = 0;
            extractAllRows('dtPindingCardMovements').forEach(cols => {
                allPendingData.push({
                    card_number:cardNum, card_name:cardName,
                    transaction_description:cols[0],
                    transaction_date:cols[1],
                    transaction_amount:cols[2]?.replace(/,/g,''),
                    debit_credit:cols[3]
                });
                pendingCount++;
            });

            // Extract POSTED
            let postedCount = 0;
            extractAllRows('dtPostedCardMovements').forEach(cols => {
                allPostedData.push({
                    card_number:cardNum, card_name:cardName, online_balance:onlineBalance,
                    transaction_description:cols[0],
                    transaction_date:cols[1],
                    transaction_amount:cols[2]?.replace(/,/g,''),
                    transaction_type:cols[3]
                });
                postedCount++;
            });

            if (postedCount === 0) {
                emptyCards.push({card_number:cardNum, card_name:cardName, online_balance:onlineBalance});
            }

            console.log(`  ✓ Posted: ${postedCount}  Pending: ${pendingCount}`);

            // Close modal and wait until fully gone before next card
            await closeModalAndWait();
            await sleep(DELAY);

        } catch(err) {
            console.error(`  ✗ Error: ${err.message}`);
            try { await closeModalAndWait(); } catch(e) {}
            await closeNotificationModal();
        }
    }

    XMLHttpRequest.prototype.open = _open;
    XMLHttpRequest.prototype.send = _send;

    const sp = sortPosted(allPostedData);
    const sn = sortPending(allPendingData);
    const sb = sortBalance(allBalanceData);

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
        link.download = `QNB_Fuel_3Sheets_${FROM_DATE.replace(/\//g,'')}_to_${TO_DATE.replace(/\//g,'')}.xls`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        console.log(`\n✅ File: ${link.download}`);
        console.log(`📊 Posted: ${sp.length}  ⏳ Pending: ${sn.length}  💰 Balance: ${sb.length}  ❌ Empty: ${emptyCards.length}`);
    } else {
        console.log('❌ No data extracted');
    }
    console.log('='.repeat(70));
})();