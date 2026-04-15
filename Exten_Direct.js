(async function extractAllFuelTransactions() {
    const FROM_DATE = '01/07/2023';
    const TO_DATE   = '30/06/2026';
    const DELAY = 300;

    const allPostedData  = [];
    const allPendingData = [];
    const allBalanceData = [];
    const emptyCards = [];
    const sleep = ms => new Promise(r => setTimeout(r, ms));

    const extractCardsFromPage = () => {
        const cards = [];
        document.querySelectorAll('#dtFuelCards tbody tr').forEach(row => {
            const cols = row.querySelectorAll('td');
            if (cols.length >= 4) {
                const cardLink = cols[3].querySelector('a.card-details');
                const refreshBtn = cols[4].querySelector('button');
                if (cardLink) {
                    const href = cardLink.getAttribute('href');
                    const match = href.match(/cardRef=(\d+)/);
                    const sigMatch = href.match(/signature=([^&]+)/);
                    const cardNum = cols[0].innerText.trim();
                    const cardName = cols[1].innerText.trim();
                    const tableBalance = cols[2].innerText.trim();
                    cards.push({
                        cardRef: match ? match[1] : '',
                        cardNumber: cardNum,
                        cardName: cardName,
                        tableBalance: tableBalance,
                        signature: sigMatch ? sigMatch[1] : ''
                    });
                }
            }
        });
        return cards;
    };

    const getOnlineBalance = async (cardRef, signature) => {
        return new Promise((resolve) => {
            $.ajax({
                url: '/FuelCard/Home/RefreshFuelCard',
                type: 'post',
                data: { CardRef: cardRef, signature: signature },
                success: function (data) {
                    if (data.Status == "0000") {
                        resolve(data.Balance);
                    } else {
                        resolve('');
                    }
                },
                error: () => resolve('')
            });
        });
    };

    const extractTableData = (tableId) => {
        const rows = [];
        try {
            if ($.fn.dataTable && $(`#${tableId}`).DataTable()) {
                $(`#${tableId}`).DataTable().rows().every(function() {
                    const node = this.node();
                    if (node) {
                        const cols = node.querySelectorAll('td');
                        if (cols.length >= 4 && !cols[0].getAttribute('colspan')) {
                            rows.push([...cols].map(c => c.innerText?.trim()));
                        }
                    }
                });
            }
        } catch(e) {
            document.getElementById(tableId)?.querySelectorAll('tbody tr').forEach(tr => {
                const cols = tr.querySelectorAll('td');
                if (cols.length >= 4 && !cols[0].getAttribute('colspan')) {
                    rows.push([...cols].map(c => c.innerText?.trim()));
                }
            });
        }
        return rows;
    };

    console.log('='.repeat(60));
    console.log('QNB FUEL EXTRACTOR - DIRECT API');
    console.log(`Date Range: ${FROM_DATE} to ${TO_DATE}`);
    console.log('='.repeat(60));

    const cards = extractCardsFromPage();
    console.log(`Found ${cards.length} cards\n`);

    for (let i = 0; i < cards.length; i++) {
        const card = cards[i];
        console.log(`[${i+1}/${cards.length}] ${card.cardNumber}`);

        try {
            const onlineBalance = await getOnlineBalance(card.cardRef, card.signature) || card.tableBalance;
            allBalanceData.push({
                card_number: card.cardNumber,
                card_name: card.cardName,
                online_balance: onlineBalance,
                table_balance: card.tableBalance,
                card_ref: card.cardRef
            });

            await new Promise(resolve => {
                $.get(`/FuelCard/Home/Details?cardRef=${card.cardRef}`, function(html) {
                    const tempDiv = document.createElement('div');
                    tempDiv.innerHTML = html;
                    
                    const postedTable = tempDiv.querySelector('#dtPostedCardMovements');
                    if (postedTable) {
                        postedTable.querySelectorAll('tbody tr').forEach(tr => {
                            const cols = tr.querySelectorAll('td');
                            if (cols.length >= 4 && !cols[0].getAttribute('colspan')) {
                                allPostedData.push({
                                    card_number: card.cardNumber,
                                    card_name: card.cardName,
                                    online_balance: onlineBalance,
                                    transaction_description: cols[0].innerText.trim(),
                                    transaction_date: cols[1].innerText.trim(),
                                    transaction_amount: cols[2].innerText.replace(/,/g, '').trim(),
                                    transaction_type: cols[3].innerText.trim()
                                });
                            }
                        });
                    }

                    const pendingTable = tempDiv.querySelector('#dtPindingCardMovements');
                    if (pendingTable) {
                        pendingTable.querySelectorAll('tbody tr').forEach(tr => {
                            const cols = tr.querySelectorAll('td');
                            if (cols.length >= 4 && !cols[0].getAttribute('colspan')) {
                                allPendingData.push({
                                    card_number: card.cardNumber,
                                    card_name: card.cardName,
                                    transaction_description: cols[0].innerText.trim(),
                                    transaction_date: cols[1].innerText.trim(),
                                    transaction_amount: cols[2].innerText.replace(/,/g, '').trim(),
                                    debit_credit: cols[3].innerText.trim()
                                });
                            }
                        });
                    }

                    const postedCount = allPostedData.filter(d => d.card_number === card.cardNumber).length;
                    if (postedCount === 0) {
                        emptyCards.push({
                            card_number: card.cardNumber,
                            card_name: card.cardName,
                            online_balance: onlineBalance
                        });
                    }

                    console.log(`  ✓ Posted: ${postedCount}`);
                    resolve();
                });
            });

            if (DELAY > 0) await sleep(DELAY);

        } catch(err) {
            console.log(`  ✗ Error: ${err.message}`);
        }
    }

    console.log(`\n✓ Posted: ${allPostedData.length}  Pending: ${allPendingData.length}  Empty: ${emptyCards.length}`);

    if (allPostedData.length || allPendingData.length || allBalanceData.length) {
        const cell = (v, s = '') => `<Cell${s ? ` ss:StyleID="${s}"` : ''}><Data ss:Type="String">${String(v ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</Data></Cell>`;
        const hcell = v => cell(v, 'h');
        const erow = (cells, i) => `<Row${i % 2 ? ' ss:StyleID="e"' : ''}>${cells}</Row>`;

        let html = '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>';
        html += '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">';
        html += '<Styles><Style ss:ID="h"><Font ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#1d2552" ss:Pattern="Solid"/></Style><Style ss:ID="b"><Font ss:Bold="1" ss:Color="#1d2552"/></Style><Style ss:ID="e"><Interior ss:Color="#f2f2f2" ss:Pattern="Solid"/></Style></Styles>';

        html += '<Worksheet ss:Name="Posted Transactions"><Table>';
        html += '<Row>' + ['Card Number', 'Card Name', 'Online Balance', 'Transaction Date', 'Transaction Description', 'Amount', 'Type (D/C)'].map(hcell).join('') + '</Row>';
        allPostedData.forEach((d, i) => html += erow(cell(d.card_number) + cell(d.card_name) + cell(d.online_balance, 'b') + cell(d.transaction_date) + cell(d.transaction_description) + cell(d.transaction_amount) + cell(d.transaction_type), i));
        html += '</Table></Worksheet>';

        html += '<Worksheet ss:Name="Pending Transactions"><Table>';
        html += '<Row>' + ['Card Number', 'Card Name', 'Transaction Date', 'Transaction Description', 'Amount', 'Debit/Credit'].map(hcell).join('') + '</Row>';
        allPendingData.forEach((d, i) => html += erow(cell(d.card_number) + cell(d.card_name) + cell(d.transaction_date) + cell(d.transaction_description) + cell(d.transaction_amount) + cell(d.debit_credit), i));
        html += '</Table></Worksheet>';

        html += '<Worksheet ss:Name="Online Balance"><Table>';
        html += '<Row>' + ['Card Number', 'Card Name', 'Online Balance', 'Table Balance', 'Card Ref'].map(hcell).join('') + '</Row>';
        allBalanceData.forEach((d, i) => html += erow(cell(d.card_number) + cell(d.card_name) + cell(d.online_balance, 'b') + cell(d.table_balance) + cell(d.card_ref), i));
        html += '</Table></Worksheet>';

        html += '<Worksheet ss:Name="Empty Cards"><Table>';
        html += '<Row>' + ['Card Number', 'Card Name', 'Online Balance'].map(hcell).join('') + '</Row>';
        emptyCards.forEach((d, i) => html += erow(cell(d.card_number) + cell(d.card_name) + cell(d.online_balance, 'b'), i));
        html += '</Table></Worksheet>';

        html += '</Workbook>';

        const blob = new Blob([html], { type: 'application/vnd.ms-excel' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `QNB_Fuel_${FROM_DATE.replace(/\//g, '')}_to_${TO_DATE.replace(/\//g, '')}.xls`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        console.log(`\n✅ File: ${link.download}`);
    }
})();