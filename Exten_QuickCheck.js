(async function findEmptyPostedCards() {
    const FROM_DATE = '01/07/2023';
    const TO_DATE   = '30/06/2026';
    const sleep = ms => new Promise(r => setTimeout(r, ms));

    const emptyCards = [];
    const cardsWithData = [];

    const waitForModalOpen = async () => {
        const start = Date.now();
        while (Date.now() - start < 8000) {
            await sleep(300);
            const m = document.getElementById('myModal');
            if (m && m.style.display !== 'none' && (m.classList.contains('in') || m.getAttribute('aria-hidden') === 'false')) {
                await sleep(500);
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
        await sleep(1500);
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
        const btn = document.querySelector('#myModal input[type="submit"], #myModal button[type="submit"], #myModal .btn-search, .modal-content input[type="submit"]');
        if (btn) { btn.click(); return true; }
        return false;
    };

    const getOnlineBalance = async (cardRef) => {
        try {
            const btn = document.querySelector(`button[onclick*="RefreshCard(${cardRef}"]`);
            if (!btn) return '';
            btn.click();
            await sleep(2500);
            const msg = document.querySelector('#notificationModal .message')?.innerText;
            const match = msg?.match(/[\d,]+\.?\d*/);
            const bal = match ? match[0] : '';
            const closeBtn = document.querySelector('#notificationModal .close, #notificationModal .notification-btn-close');
            if (closeBtn) closeBtn.click();
            await sleep(500);
            return bal;
        } catch(e) { return ''; }
    };

    const checkPostedCount = () => {
        try {
            if (window.jQuery && $.fn.dataTable?.isDataTable('#dtPostedCardMovements')) {
                return $('#dtPostedCardMovements').DataTable().page.info().recordsTotal;
            }
        } catch(e) {}
        const rows = document.getElementById('dtPostedCardMovements')?.querySelectorAll('tbody tr') || [];
        return (rows.length === 1 && rows[0].querySelector('td[colspan]')) ? 0 : rows.length;
    };

    console.log('='.repeat(60));
    console.log('FAST EMPTY CARDS CHECKER (with Balance)');
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

        console.log(`[${i+1}/${cards.length}] ${cardNum}...`);

        try {
            const onlineBalance = await getOnlineBalance(cardRef) || tableBalance;

            card.click();
            await waitForModalOpen();

            await setDate('From', FROM_DATE);
            await sleep(100);
            await setDate('To', TO_DATE);
            await sleep(100);

            clickSearch();
            await sleep(3000);

            const count = checkPostedCount();

            if (count === 0) {
                console.log(`  ❌ EMPTY | Balance: ${onlineBalance}`);
                emptyCards.push({card_number: cardNum, card_name: cardName, online_balance: onlineBalance});
            } else {
                console.log(`  ✅ ${count} txns | Balance: ${onlineBalance}`);
                cardsWithData.push({card_number: cardNum, card_name: cardName, online_balance: onlineBalance, count: count});
            }

            await closeModalAndWait();
            await sleep(800);

        } catch(err) {
            console.log(`  ⚠ Error: ${err.message}`);
            try { await closeModalAndWait(); } catch(e) {}
        }
    }

    console.log('\n' + '='.repeat(60));
    console.log('RESULTS:');
    console.log(`  Empty cards: ${emptyCards.length}`);
    console.log(`  Cards with data: ${cardsWithData.length}`);
    console.log('='.repeat(60));

    if (emptyCards.length > 0) {
        console.log('\n📋 EMPTY CARDS:');
        emptyCards.forEach(c => console.log(`  ${c.card_number} | ${c.card_name} | Balance: ${c.online_balance}`));
    }

    const escapeXML = s => String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

    const html = '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"><Styles><Style ss:ID="h"><Font ss:Bold="1"/><Interior ss:Color="#1d2552" ss:Pattern="Solid"/></Style></Styles>' +
        '<Worksheet ss:Name="Empty Cards"><Table><Row><Cell ss:StyleID="h"><Data ss:Type="String">Card Number</Data></Cell><Cell ss:StyleID="h"><Data ss:Type="String">Card Name</Data></Cell><Cell ss:StyleID="h"><Data ss:Type="String">Online Balance</Data></Cell></Row>' + 
        emptyCards.map(c => `<Row><Cell><Data ss:Type="String">${escapeXML(c.card_number)}</Data></Cell><Cell><Data ss:Type="String">${escapeXML(c.card_name)}</Data></Cell><Cell><Data ss:Type="String">${escapeXML(c.online_balance)}</Data></Cell></Row>`).join('') + 
        '</Table></Worksheet>' +
        '<Worksheet ss:Name="Cards With Data"><Table><Row><Cell ss:StyleID="h"><Data ss:Type="String">Card Number</Data></Cell><Cell ss:StyleID="h"><Data ss:Type="String">Card Name</Data></Cell><Cell ss:StyleID="h"><Data ss:Type="String">Online Balance</Data></Cell><Cell ss:StyleID="h"><Data ss:Type="String">Transactions</Data></Cell></Row>' + 
        cardsWithData.map(c => `<Row><Cell><Data ss:Type="String">${escapeXML(c.card_number)}</Data></Cell><Cell><Data ss:Type="String">${escapeXML(c.card_name)}</Data></Cell><Cell><Data ss:Type="String">${escapeXML(c.online_balance)}</Data></Cell><Cell><Data ss:Type="String">${c.count}</Data></Cell></Row>`).join('') + 
        '</Table></Worksheet></Workbook>';

    const blob = new Blob([html], {type:'application/vnd.ms-excel'});
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `Cards_Check_${new Date().toISOString().slice(0,10)}.xls`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    console.log(`\n✅ Saved: ${link.download}`);
})();