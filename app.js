
const App = (() => {

    /* =========================================
       APPLICATION STATE
    ========================================== */

    const state = {

        rows: [],

        brokers: [],

        clients: [],

        source: "",

        loadedAt: null,

        chart: null

    };


    /* =========================================
       FORMATTERS
    ========================================== */

    const money = new Intl.NumberFormat("en-US", {
        style: "currency",
        currency: "USD",
        minimumFractionDigits: 2
    });


    const number = new Intl.NumberFormat("en-US");


    /* =========================================
       DOM HELPER
    ========================================== */

    const $ = id => document.getElementById(id);


    /* =========================================
       TEXT NORMALIZATION
    ========================================== */

    const norm = value => {

        return String(value ?? "")
            .trim()
            .replace(/\s+/g, " ")
            .toLowerCase();

    };


    const title = value => {

        return String(value ?? "")
            .replace(/\s+/g, " ")
            .trim()
            .replace(/\b\w/g, character =>
                character.toUpperCase()
            );

    };


    /* =========================================
       MONEY HELPERS
    ========================================== */

    const cents = value => {

        return Math.round(
            Number(value || 0) * 100
        );

    };


    const fmt = value => {

        return money.format(
            Number(value || 0)
        );

    };


    /* =========================================
       HTML ESCAPE
    ========================================== */

    const esc = value => {

        return String(value ?? "").replace(
            /[&<>"']/g,

            character => ({

                "&": "&amp;",
                "<": "&lt;",
                ">": "&gt;",
                '"': "&quot;",
                "'": "&#39;"

            })[character]
        );

    };


    /* =========================================
       TOAST
    ========================================== */

    const toast = message => {

        const element = $("toast");

        element.textContent = message;

        element.classList.add("show");

        setTimeout(() => {

            element.classList.remove("show");

        }, 2600);

    };


    /* =========================================
       DATE PARSING
    ========================================== */

    function parseDate(value) {

        if (
            value instanceof Date &&
            !isNaN(value)
        ) {

            return value;

        }


        /*
         * Excel serial date
         */

        if (typeof value === "number") {

            const parsed =
                XLSX.SSF.parse_date_code(value);

            return parsed
                ? new Date(
                    parsed.y,
                    parsed.m - 1,
                    parsed.d
                )
                : null;

        }


        const stringValue =
            String(value || "").trim();


        if (!stringValue) {

            return null;

        }


        /*
         * MM/DD/YYYY
         * MM-DD-YYYY
         */

        const match =
            stringValue.match(
                /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/
            );


        if (match) {

            return new Date(
                +match[3],
                +match[1] - 1,
                +match[2]
            );

        }


        const date =
            new Date(stringValue);


        return isNaN(date)
            ? null
            : date;

    }


    /* =========================================
       DATE LABEL
    ========================================== */

    function dateLabel(date) {

        if (!date) {

            return "—";

        }


        return new Intl.DateTimeFormat(
            "en-US",
            {
                month: "2-digit",
                day: "2-digit",
                year: "numeric"
            }
        ).format(date);

    }


    /* =========================================
       AGE
    ========================================== */

    function ageDays(date) {

        if (!date) {

            return null;

        }


        return Math.max(
            0,
            Math.floor(
                (
                    Date.now() -
                    date.getTime()
                ) /
                86400000
            )
        );

    }


    /* =========================================
       FIND REPORT HEADER
    ========================================== */

    function findHeader(matrix) {

        const wanted = [

            "transaction date",

            "transaction type",

            "num",

            "name",

            "cliente",

            "broker",

            "amount",

            "balance"

        ];


        for (
            let i = 0;
            i < Math.min(matrix.length, 30);
            i++
        ) {

            const cells =
                (matrix[i] || [])
                    .map(value => norm(value));


            const hits =
                wanted.filter(
                    column =>
                        cells.includes(column)
                ).length;


            /*
             * The report is considered valid
             * if at least 5 expected headers
             * are found.
             */

            if (hits >= 5) {

                return {

                    row: i,

                    cells

                };

            }

        }


        throw new Error(
            "Could not find the expected transaction headers."
        );

    }


    /* =========================================
       PARSE WORKBOOK
    ========================================== */

    function parseWorkbook(
        buffer,
        source = "Uploaded report"
    ) {

        const workbook =
            XLSX.read(
                buffer,
                {
                    type: "array",
                    cellDates: true
                }
            );


        /*
         * Use the first worksheet.
         */

        const worksheet =
            workbook.Sheets[
                workbook.SheetNames[0]
            ];


        /*
         * Convert worksheet to matrix.
         */

        const matrix =
            XLSX.utils.sheet_to_json(
                worksheet,
                {
                    header: 1,
                    raw: true,
                    defval: null
                }
            );


        /*
         * Find header row.
         */

        const header =
            findHeader(matrix);


        /*
         * Create header → column index map.
         */

        const map = {};


        header.cells.forEach(
            (headerName, index) => {

                if (headerName) {

                    map[headerName] = index;

                }

            }
        );


        const rows = [];


        /*
         * Process transaction rows.
         */

        for (
            let i = header.row + 1;
            i < matrix.length;
            i++
        ) {

            const row =
                matrix[i] || [];


            const date =
                parseDate(
                    row[map["transaction date"]]
                );


            const type =
                row[map["transaction type"]];


            const num =
                row[map["num"]];


            const client =
                row[map["cliente"]];


            const broker =
                row[map["broker"]];


            const amount =
                Number(
                    row[map["amount"]]
                );


            const balance =
                Number(
                    row[map["balance"]]
                );


            /*
             * Ignore invalid rows.
             */

            if (
                !date ||
                !type ||
                !broker ||
                !Number.isFinite(amount)
            ) {

                continue;

            }


            rows.push({

                id:
                    `${i}-${num}-${amount}`,

                date,

                type:
                    String(type).trim(),

                num:
                    String(num ?? "").trim(),

                client:
                    String(client ?? "").trim(),

                broker:
                    String(broker).trim(),

                brokerKey:
                    norm(broker),

                amount,

                balance:
                    Number.isFinite(balance)
                        ? balance
                        : null,

                age:
                    ageDays(date)

            });

        }


        if (!rows.length) {

            throw new Error(
                "No usable transaction rows were found in the report."
            );

        }


        return {

            rows,

            source,

            loadedAt:
                new Date()

        };

    }


    /* =========================================
       LOCAL STORAGE
    ========================================== */

    async function saveLocal() {

        try {

            localStorage.setItem(

                "wc_income_rows",

                JSON.stringify({

                    rows:
                        state.rows.map(row => ({

                            ...row,

                            date:
                                row.date.toISOString()

                        })),

                    source:
                        state.source,

                    loadedAt:
                        state.loadedAt.toISOString()

                })

            );

        } catch (error) {

            console.warn(
                "Local save failed",
                error
            );

        }

    }


    function restoreLocal() {

        try {

            const raw =
                localStorage.getItem(
                    "wc_income_rows"
                );


            if (!raw) {

                return false;

            }


            const saved =
                JSON.parse(raw);


            state.rows =
                saved.rows.map(row => ({

                    ...row,

                    date:
                        new Date(row.date)

                }));


            state.source =
                saved.source;


            state.loadedAt =
                new Date(
                    saved.loadedAt
                );


            return true;

        } catch (error) {

            return false;

        }

    }


    /* =========================================
       APPLY DATASET
    ========================================== */

    function applyDataset(parsed) {

        state.rows =
            parsed.rows;

        state.source =
            parsed.source;

        state.loadedAt =
            parsed.loadedAt;


        /*
         * Build broker list.
         */

        state.brokers = [

            ...new Map(

                state.rows.map(row => [

                    row.brokerKey,

                    title(row.broker)

                ])

            ).entries()

        ]

            .map(
                ([key, name]) => ({
                    key,
                    name
                })
            )

            .sort(
                (a, b) =>
                    a.name.localeCompare(
                        b.name
                    )
            );


        /*
         * Build client list.
         */

        state.clients = [

            ...new Set(

                state.rows

                    .map(row => row.client)

                    .filter(Boolean)

                    .map(title)

            )

        ].sort();


        /*
         * Update sidebar.
         */

        $("sidebarStatus").textContent =
            `${number.format(
                state.rows.length
            )} rows ready`;


        /*
         * Update report date.
         */

        $("reportDate").textContent =
            `Updated ${
                new Intl.DateTimeFormat(
                    "en-US",
                    {
                        month: "short",
                        day: "2-digit",
                        year: "numeric",
                        hour: "numeric",
                        minute: "2-digit"
                    }
                ).format(state.loadedAt)
            }`;


        /*
         * Populate broker autocomplete.
         */

        const list =
            $("brokerList");


        list.innerHTML =
            state.brokers
                .map(
                    broker =>
                        `<option value="${esc(
                            broker.name
                        )}"></option>`
                )
                .join("");


        /*
         * Refresh UI.
         */

        renderDashboard();

        renderBrokers();

        renderData();

    }


    /* =========================================
       LOAD FILE
    ========================================== */

    async function loadFile(file) {

        if (!file) {

            return;

        }


        try {

            const parsed =
                parseWorkbook(
                    await file.arrayBuffer(),
                    file.name
                );


            applyDataset(parsed);


            await saveLocal();


            toast(
                `Loaded ${
                    number.format(
                        parsed.rows.length
                    )
                } transactions.`
            );

        } catch (error) {

            console.error(error);

            toast(error.message);

        }

    }


    /* =========================================
       AUTO LOAD
    ========================================== */

    async function tryAutoLoad() {

        /*
         * First try local browser data.
         */

        if (restoreLocal()) {

            applyDataset({

                rows:
                    state.rows,

                source:
                    state.source,

                loadedAt:
                    state.loadedAt

            });

            return;

        }


        /*
         * If there is no local dataset,
         * try the default Excel file.
         */

        try {

            const response =
                await fetch(
                    "data/AR_Report.xlsx",
                    {
                        cache: "no-store"
                    }
                );


            if (!response.ok) {

                throw new Error();

            }


            const parsed =
                parseWorkbook(
                    await response.arrayBuffer(),
                    "data/AR_Report.xlsx"
                );


            applyDataset(parsed);


            await saveLocal();

        } catch (error) {

            $("sidebarStatus").textContent =
                "Waiting for report";


            $("dataDescription").textContent =
                "No report loaded. Use Data → Choose Excel report.";

        }

    }


    /* =========================================
       DASHBOARD
    ========================================== */

    function renderDashboard() {

        const total =
            state.rows.reduce(
                (sum, row) =>
                    sum + row.amount,
                0
            );


        $("kpiMovements").textContent =
            number.format(
                state.rows.length
            );


        $("kpiAmount").textContent =
            fmt(total);


        $("kpiBrokers").textContent =
            number.format(
                state.brokers.length
            );


        $("kpiClients").textContent =
            number.format(
                state.clients.length
            );


        /*
         * Top 10 brokers.
         */

        const groups =
            aggregateBrokers()
                .slice(0, 10);


        /*
         * Destroy previous chart.
         */

        const canvas =
            $("brokerChart");


        if (state.chart) {

            state.chart.destroy();

        }


        /*
         * Create chart.
         */

        state.chart =
            new Chart(
                canvas,
                {

                    type: "bar",

                    data: {

                        labels:
                            groups.map(
                                broker =>
                                    shorten(
                                        broker.name,
                                        23
                                    )
                            ),

                        datasets: [

                            {

                                data:
                                    groups.map(
                                        broker =>
                                            broker.amount
                                    ),

                                backgroundColor:
                                    "#6d4aff",

                                borderRadius: 5,

                                barThickness: 18

                            }

                        ]

                    },

                    options: {

                        indexAxis: "y",

                        plugins: {

                            legend: {
                                display: false
                            },

                            tooltip: {

                                callbacks: {

                                    label: context =>
                                        fmt(
                                            context.raw
                                        )

                                }

                            }

                        },

                        scales: {

                            x: {

                                grid: {
                                    color: "#eef0f4"
                                },

                                ticks: {

                                    font: {
                                        size: 9
                                    },

                                    callback: value =>
                                        fmt(value)

                                }

                            },

                            y: {

                                grid: {
                                    display: false
                                },

                                ticks: {

                                    font: {
                                        size: 9
                                    }

                                }

                            }

                        }

                    }

                }
            );


        /*
         * Top brokers table.
         */

        $("topBrokersTable").innerHTML = `

            <table class="mini-table">

                <thead>

                    <tr>

                        <th>
                            Broker
                        </th>

                        <th>
                            Items
                        </th>

                        <th>
                            Amount
                        </th>

                    </tr>

                </thead>

                <tbody>

                    ${groups.map(broker => `

                        <tr>

                            <td>
                                ${esc(broker.name)}
                            </td>

                            <td>
                                ${number.format(
                                    broker.count
                                )}
                            </td>

                            <td>
                                ${fmt(
                                    broker.amount
                                )}
                            </td>

                        </tr>

                    `).join("")}

                </tbody>

            </table>

        `;

    }


    /* =========================================
       BROKER AGGREGATION
    ========================================== */

    function aggregateBrokers() {

        const map =
            new Map();


        for (const row of state.rows) {

            const existing =
                map.get(
                    row.brokerKey
                ) || {

                    key:
                        row.brokerKey,

                    name:
                        title(row.broker),

                    amount: 0,

                    count: 0,

                    clients:
                        new Set(),

                    oldest:
                        null

                };


            existing.amount +=
                row.amount;


            existing.count++;


            if (row.client) {

                existing.clients.add(
                    norm(row.client)
                );

            }


            if (
                !existing.oldest ||
                row.date < existing.oldest
            ) {

                existing.oldest =
                    row.date;

            }


            map.set(
                row.brokerKey,
                existing
            );

        }


        return [
            ...map.values()
        ].sort(
            (a, b) =>
                b.amount - a.amount
        );

    }


    /* =========================================
       SHORTEN TEXT
    ========================================== */

    function shorten(
        value,
        maxLength
    ) {

        return value.length > maxLength

            ? value.slice(
                0,
                maxLength - 1
            ) + "…"

            : value;

    }


    /* =========================================
       BROKER LIST
    ========================================== */

    function renderBrokers(filter = "") {

        const query =
            norm(filter);


        const all =
            aggregateBrokers();


        const list =
            all.filter(
                broker =>
                    !query ||
                    norm(
                        broker.name
                    ).includes(query)
            );


        $("brokerCount").textContent =
            `${number.format(
                list.length
            )} brokers`;


        $("brokerCards").innerHTML =

            list.map(broker => `

                <article
                    class="broker-card"
                    data-broker="${esc(
                        broker.key
                    )}"
                >

                    <h3
                        title="${esc(
                            broker.name
                        )}"
                    >
                        ${esc(
                            broker.name
                        )}
                    </h3>


                    <p>

                        ${number.format(
                            broker.clients.size
                        )}
                        clients

                        ·

                        ${number.format(
                            broker.count
                        )}
                        movements

                    </p>


                    <div class="amount">

                        ${fmt(
                            broker.amount
                        )}

                    </div>


                    <div class="broker-meta">

                        <span>

                            Oldest
                            ${dateLabel(
                                broker.oldest
                            )}

                        </span>

                        <span>
                            View →
                        </span>

                    </div>

                </article>

            `).join("")


            ||

            `

                <div class="empty-state">

                    <h2>
                        No brokers found
                    </h2>

                    <p>
                        Try another search.
                    </p>

                </div>

            `;


        /*
         * Attach click events.
         */

        document
            .querySelectorAll(".broker-card")
            .forEach(card => {

                card.addEventListener(
                    "click",
                    () =>
                        openBroker(
                            card.dataset.broker
                        )
                );

            });

    }


    /* =========================================
       IDENTIFICATION
    ========================================== */

    function identify() {

        const broker =
            norm(
                $("brokerInput").value
            );


        const amount =
            Number(
                String(
                    $("amountInput").value
                )
                    .replace(
                        /[$,\s]/g,
                        ""
                    )
            );


        const date =
            $("dateInput").value

                ? new Date(
                    $("dateInput").value +
                    "T00:00:00"
                )

                : null;


        /*
         * Validate.
         */

        if (
            !broker ||
            !Number.isFinite(amount) ||
            amount <= 0
        ) {

            toast(
                "Enter a broker and a valid payment amount."
            );

            return;

        }


        /*
         * Find movements belonging
         * to the broker.
         */

        const candidates =
            state.rows

                .filter(
                    row =>
                        row.brokerKey === broker ||
                        norm(
                            row.broker
                        ).includes(broker)
                )

                .filter(
                    row =>
                        row.amount > 0
                );


        if (!candidates.length) {

            renderNoMatch();

            return;

        }


        /*
         * Exact individual matches.
         */

        const exact =
            candidates

                .filter(
                    row =>
                        cents(row.amount) ===
                        cents(amount)
                )

                .sort(
                    (a, b) => {

                        if (!date) {

                            return 0;

                        }


                        return (
                            Math.abs(
                                a.date - date
                            ) -
                            Math.abs(
                                b.date - date
                            )
                        );

                    }
                );


        /*
         * Combination matching.
         */

        const combo =
            findCombination(
                candidates,
                amount,
                date
            );


        /*
         * Render.
         */

        renderIdentification(
            broker,
            amount,
            date,
            exact,
            combo,
            candidates
        );

    }


    /* =========================================
       COMBINATION MATCHING
    ========================================== */

    function findCombination(
        rows,
        target,
        date
    ) {

        const targetC =
            cents(target);


        /*
         * Ignore movements above
         * the target payment.
         */

        let candidates =
            [...rows]

                .filter(
                    row =>
                        cents(row.amount) <=
                        targetC
                )

                .sort(
                    (a, b) => {

                        const dateA =
                            date
                                ? Math.abs(
                                    a.date - date
                                )
                                : 0;


                        const dateB =
                            date
                                ? Math.abs(
                                    b.date - date
                                )
                                : 0;


                        return (
                            (b.amount - a.amount) +
                            (
                                (dateA - dateB) /
                                86400000
                            ) *
                            0.0001
                        );

                    }
                )

                /*
                 * Limit search size
                 * to protect browser performance.
                 */

                .slice(0, 120);


        /*
         * Dynamic programming map.
         *
         * key   = accumulated cents
         * value = rows that create the amount
         */

        const dp =
            new Map([
                [0, []]
            ]);


        let best = null;


        for (
            const row of candidates
        ) {

            const rowCents =
                cents(row.amount);


            const entries =
                [...dp.entries()];


            for (
                const [sum, combination]
                of entries
            ) {

                const newSum =
                    sum + rowCents;


                if (
                    newSum > targetC
                ) {

                    continue;

                }


                /*
                 * Only save first combination
                 * for each accumulated amount.
                 */

                if (
                    !dp.has(newSum)
                ) {

                    const newCombination =
                        [
                            ...combination,
                            row
                        ];


                    dp.set(
                        newSum,
                        newCombination
                    );


                    /*
                     * Exact combination found.
                     */

                    if (
                        newSum === targetC
                    ) {

                        best =
                            newCombination;

                        break;

                    }

                }

            }


            if (best) {

                break;

            }


            /*
             * Prevent excessive browser memory.
             */

            if (
                dp.size > 100000
            ) {

                break;

            }

        }


        return best;

    }


    /* =========================================
       NO MATCH
    ========================================== */

    function renderNoMatch() {

        $("identifyEmpty")
            .classList
            .add("hidden");


        const box =
            $("identifyResults");


        box.classList.remove(
            "hidden"
        );


        box.innerHTML = `

            <div class="result-head">

                <div>

                    <h2>
                        No candidate found
                    </h2>

                    <p>
                        No movement for that broker
                        could be matched to the
                        requested amount.
                    </p>

                </div>

                <span class="badge red">
                    NO MATCH
                </span>

            </div>

        `;

    }


    /* =========================================
       IDENTIFICATION RESULTS
    ========================================== */

    function renderIdentification(
        broker,
        amount,
        date,
        exact,
        combo,
        candidates
    ) {

        $("identifyEmpty")
            .classList
            .add("hidden");


        const box =
            $("identifyResults");


        box.classList.remove(
            "hidden"
        );


        let html = `

            <div class="result-head">

                <div>

                    <h2>
                        Identification result
                    </h2>

                    <p>

                        ${number.format(
                            candidates.length
                        )}

                        movements found for

                        <strong>
                            ${esc(
                                title(broker)
                            )}
                        </strong>.

                    </p>

                </div>

            </div>

        `;


        /*
         * Exact individual matches.
         */

        if (exact.length) {

            html +=
                matchCard(
                    "Exact amount match",
                    exact.slice(0, 10),
                    amount,
                    date,
                    "green",
                    "EXACT MATCH"
                );

        }


        /*
         * Exact combination.
         */

        if (
            combo &&
            (
                !exact.length ||
                combo.length > 1
            )
        ) {

            html +=
                matchCard(
                    "Possible batch match",
                    combo,
                    amount,
                    date,
                    "green",
                    "EXACT COMBINATION"
                );

        }


        /*
         * Closest candidates.
         */

        if (
            !exact.length &&
            !combo
        ) {

            const nearest =
                [...candidates]

                    .sort(
                        (a, b) =>
                            Math.abs(
                                a.amount -
                                amount
                            ) -
                            Math.abs(
                                b.amount -
                                amount
                            )
                    )

                    .slice(0, 10);


            html +=
                matchCard(
                    "Closest open movements",
                    nearest,
                    amount,
                    date,
                    "yellow",
                    "REVIEW"
                );

        }


        box.innerHTML =
            html;

    }


    /* =========================================
       MATCH CARD
    ========================================== */

    function matchCard(
        titleText,
        rows,
        payment,
        date,
        badgeClass,
        badgeText
    ) {

        const sum =
            rows.reduce(
                (total, row) =>
                    total + row.amount,
                0
            );


        const difference =
            payment - sum;


        return `

            <article class="match-card">


                <!-- Header -->

                <div class="match-top">

                    <div class="match-title">

                        <h3>
                            ${esc(
                                titleText
                            )}
                        </h3>

                        <p>

                            ${rows.length}

                            movement
                            ${rows.length === 1
                                ? ""
                                : "s"}

                            ·

                            ${
                                date
                                    ? `payment date ${dateLabel(
                                        date
                                    )}`
                                    : "no payment date provided"
                            }

                        </p>

                    </div>


                    <span
                        class="badge ${badgeClass}"
                    >
                        ${badgeText}
                    </span>

                </div>


                <!-- Summary -->

                <div class="match-summary">


                    <div class="summary-item">

                        <small>
                            Selected
                        </small>

                        <strong>
                            ${fmt(sum)}
                        </strong>

                    </div>


                    <div class="summary-item">

                        <small>
                            Payment
                        </small>

                        <strong>
                            ${fmt(payment)}
                        </strong>

                    </div>


                    <div class="summary-item">

                        <small>
                            Difference
                        </small>

                        <strong>
                            ${fmt(difference)}
                        </strong>

                    </div>


                </div>


                <!-- Movements -->

                <table class="match-table">

                    <thead>

                        <tr>

                            <th>
                                Date
                            </th>

                            <th>
                                Num
                            </th>

                            <th>
                                Client
                            </th>

                            <th>
                                Broker
                            </th>

                            <th>
                                Amount
                            </th>

                        </tr>

                    </thead>


                    <tbody>

                        ${rows.map(row => `

                            <tr>

                                <td>
                                    ${dateLabel(
                                        row.date
                                    )}
                                </td>

                                <td>
                                    ${esc(
                                        row.num
                                    )}
                                </td>

                                <td>
                                    ${esc(
                                        title(
                                            row.client
                                        )
                                    )}
                                </td>

                                <td>
                                    ${esc(
                                        title(
                                            row.broker
                                        )
                                    )}
                                </td>

                                <td class="amount">

                                    ${fmt(
                                        row.amount
                                    )}

                                </td>

                            </tr>

                        `).join("")}

                    </tbody>

                </table>

            </article>

        `;

    }


    /* =========================================
       BROKER DETAIL MODAL
    ========================================== */

    function openBroker(key) {

        const rows =
            state.rows

                .filter(
                    row =>
                        row.brokerKey === key
                )

                .sort(
                    (a, b) =>
                        b.date - a.date
                );


        if (!rows.length) {

            return;

        }


        const name =
            title(
                rows[0].broker
            );


        const total =
            rows.reduce(
                (sum, row) =>
                    sum + row.amount,
                0
            );


        const clients =
            new Set(
                rows.map(
                    row =>
                        norm(row.client)
                )
            );


        $("modalContent").innerHTML = `

            <div class="detail-title">

                <h2>
                    ${esc(name)}
                </h2>

                <p>
                    Movement history available
                    in the current AR report.
                </p>

            </div>


            <div class="detail-stats">


                <div class="detail-stat">

                    <small>
                        Total amount
                    </small>

                    <strong>
                        ${fmt(total)}
                    </strong>

                </div>


                <div class="detail-stat">

                    <small>
                        Movements
                    </small>

                    <strong>
                        ${number.format(
                            rows.length
                        )}
                    </strong>

                </div>


                <div class="detail-stat">

                    <small>
                        Clients
                    </small>

                    <strong>
                        ${number.format(
                            clients.size
                        )}
                    </strong>

                </div>


            </div>


            <table class="detail-table">

                <thead>

                    <tr>

                        <th>
                            Date
                        </th>

                        <th>
                            Num
                        </th>

                        <th>
                            Client
                        </th>

                        <th>
                            Type
                        </th>

                        <th>
                            Amount
                        </th>

                        <th>
                            Balance
                        </th>

                    </tr>

                </thead>


                <tbody>

                    ${rows.map(row => `

                        <tr>

                            <td>
                                ${dateLabel(
                                    row.date
                                )}
                            </td>

                            <td>
                                ${esc(
                                    row.num
                                )}
                            </td>

                            <td>
                                ${esc(
                                    title(
                                        row.client
                                    )
                                )}
                            </td>

                            <td>
                                ${esc(
                                    row.type
                                )}
                            </td>

                            <td>
                                ${fmt(
                                    row.amount
                                )}
                            </td>

                            <td>
                                ${
                                    row.balance == null
                                        ? "—"
                                        : fmt(
                                            row.balance
                                        )
                                }
                            </td>

                        </tr>

                    `).join("")}

                </tbody>

            </table>

        `;


        $("modal")
            .classList
            .remove("hidden");

    }


    /* =========================================
       DATA VIEW
    ========================================== */

    function renderData() {

        const total =
            state.rows.reduce(
                (sum, row) =>
                    sum + row.amount,
                0
            );


        $("dataDescription").textContent =

            `${state.source} · loaded ${
                new Intl.DateTimeFormat(
                    "en-US",
                    {
                        dateStyle: "medium",
                        timeStyle: "short"
                    }
                ).format(
                    state.loadedAt
                )
            }`;


        const dates =
            state.rows
                .map(row =>
                    row.date.getTime()
                );


        const minimumDate =
            Math.min(...dates);


        const maximumDate =
            Math.max(...dates);


        $("dataStats").innerHTML = [

            [
                "Rows",
                number.format(
                    state.rows.length
                )
            ],

            [
                "Brokers",
                number.format(
                    state.brokers.length
                )
            ],

            [
                "Clients",
                number.format(
                    state.clients.length
                )
            ],

            [
                "Amount",
                fmt(total)
            ],

            [
                "First date",
                dateLabel(
                    Number.isFinite(
                        minimumDate
                    )
                        ? new Date(
                            minimumDate
                        )
                        : null
                )
            ],

            [
                "Last date",
                dateLabel(
                    Number.isFinite(
                        maximumDate
                    )
                        ? new Date(
                            maximumDate
                        )
                        : null
                )
            ]

        ]

            .map(
                ([label, value]) => `

                    <div class="data-stat">

                        <span>
                            ${label}
                        </span>

                        <strong>
                            ${value}
                        </strong>

                    </div>

                `
            )

            .join("");

    }


    /* =========================================
       SWITCH VIEW
    ========================================== */

    function switchView(view) {

        /*
         * Update navigation.
         */

        document
            .querySelectorAll(".nav-item")
            .forEach(button => {

                button.classList.toggle(
                    "active",
                    button.dataset.view === view
                );

            });


        /*
         * Hide every view.
         */

        document
            .querySelectorAll(".view")
            .forEach(section => {

                section.classList.remove(
                    "active-view"
                );

            });


        /*
         * Display selected view.
         */

        $(`view-${view}`)
            .classList
            .add("active-view");

    }


    /* =========================================
       GLOBAL SEARCH
    ========================================== */

    function globalSearch(query) {

        query =
            norm(query);


        if (!query) {

            return;

        }


        /*
         * Search broker first.
         */

        const broker =
            state.brokers.find(
                item =>
                    norm(
                        item.name
                    ).includes(query)
            );


        if (broker) {

            switchView(
                "brokers"
            );


            $("brokerFilter").value =
                broker.name;


            renderBrokers(
                broker.name
            );


            return;

        }


        /*
         * Search client.
         */

        const client =
            state.rows.find(
                row =>
                    norm(
                        row.client
                    ).includes(query)
            );


        if (client) {

            $("brokerInput").value =
                client.broker;


            $("amountInput").value =
                client.amount;


            switchView(
                "identify"
            );


            identify();


            return;

        }


        /*
         * Search transaction number.
         */

        const transaction =
            state.rows.find(
                row =>
                    norm(
                        row.num
                    ) === query
            );


        if (transaction) {

            $("brokerInput").value =
                transaction.broker;


            $("amountInput").value =
                transaction.amount;


            switchView(
                "identify"
            );


            identify();


            return;

        }


        /*
         * Search amount.
         */

        const amount =
            Number(
                query.replace(
                    /[$,\s]/g,
                    ""
                )
            );


        if (
            Number.isFinite(amount)
        ) {

            switchView(
                "identify"
            );


            $("amountInput").value =
                amount;


            toast(
                "Enter the broker to narrow the identification."
            );

        }

    }


    /* =========================================
       EVENT BINDINGS
    ========================================== */

    function bind() {


        /*
         * Navigation.
         */

        document
            .querySelectorAll(".nav-item")
            .forEach(button => {

                button.addEventListener(
                    "click",
                    () =>
                        switchView(
                            button.dataset.view
                        )
                );

            });


        /*
         * Identification.
         */

        $("identifyBtn")
            .addEventListener(
                "click",
                identify
            );


        $("amountInput")
            .addEventListener(
                "keydown",
                event => {

                    if (
                        event.key === "Enter"
                    ) {

                        identify();

                    }

                }
            );


        /*
         * Broker search.
         */

        $("brokerFilter")
            .addEventListener(
                "input",
                event =>
                    renderBrokers(
                        event.target.value
                    )
            );


        /*
         * Global search.
         */

        $("globalSearch")
            .addEventListener(
                "keydown",
                event => {

                    if (
                        event.key === "Enter"
                    ) {

                        globalSearch(
                            event.target.value
                        );

                    }

                }
            );


        /*
         * File input.
         */

        $("fileInput")
            .addEventListener(
                "change",
                event =>
                    loadFile(
                        event.target.files[0]
                    )
            );


        /*
         * Reload report.
         */

        $("refreshBtn")
            .addEventListener(
                "click",
                tryAutoLoad
            );


        /*
         * Close modal.
         */

        $("modalClose")
            .addEventListener(
                "click",
                () =>
                    $("modal")
                        .classList
                        .add("hidden")
            );


        /*
         * Close modal clicking backdrop.
         */

        document
            .querySelector(".modal-backdrop")
            .addEventListener(
                "click",
                () =>
                    $("modal")
                        .classList
                        .add("hidden")
            );


        /*
         * Drag and drop.
         */

        const dropZone =
            $("dropZone");


        [
            "dragenter",
            "dragover"
        ].forEach(eventName => {

            dropZone.addEventListener(
                eventName,
                event => {

                    event.preventDefault();

                    dropZone.style.borderColor =
                        "#6d4aff";

                }
            );

        });


        [
            "dragleave",
            "drop"
        ].forEach(eventName => {

            dropZone.addEventListener(
                eventName,
                event => {

                    event.preventDefault();

                    dropZone.style.borderColor =
                        "";

                }
            );

        });


        dropZone.addEventListener(
            "drop",
            event =>
                loadFile(
                    event.dataTransfer.files[0]
                )
        );

    }


    /* =========================================
       APPLICATION START
    ========================================== */

    bind();

    tryAutoLoad();


    /* =========================================
       PUBLIC API
    ========================================== */

    return {

        loadFile

    };

})();


