const dashboardContent = `
    <div> 
  </div> <br>
    <div class="kpi-grid">
        <!-- KPI 1 -->
        <div class="kpi-card">
            <h3>Order Backlog (OBL)</h3>
            <div class="kpi-value blue">9,140,362</div>
            <div class="kpi-change red">(-2.84%)</div>
        </div>

        <!-- KPI 2 Wide2 -->
        <div class="kpi-card wide2">
            <div class="dflex split">
                <div class="kpi-left">
                    <h3>Total Receivables</h3>
                    <div class="kpi-value blue">4,127,071</div>
                    <div class="kpi-change red">(+18.64%)</div>
                </div>
                <div class="kpi-right">
                    <div>
                        <div class="kpi-card-box">
                            <div class="kpi-percent red">(+74.93%)</div>
                            <div class="kpi-value-2 blue">1,444,466</div>
                            <h5 class="kpi-lable">Normal</h5>
                        </div>
                        <div class="kpi-card-box">
                            <div class="kpi-percent red">(+26.54%)</div>
                            <div class="kpi-value-2 blue">669,655</div>
                            <h5 class="kpi-lable">Retention</h5>
                        </div>
                    </div>
                    <div>
                        <div class="kpi-card-box">
                            <div class="kpi-percent green">(+11.93%)</div>
                            <div class="kpi-value-2 blue">1,023,367</div>
                            <h5 class="kpi-lable">Normal Overdue</h5>
                        </div>
                        <div class="kpi-card-box">
                            <div class="kpi-percent red">(+1.67%)</div>
                            <div class="kpi-value-2 blue">866,921</div>
                            <h5 class="kpi-lable">Retention Overdue</h5>
                        </div>
                    </div>
                    <div>
                        <div class="kpi-card-box">
                            <div class="kpi-percent red">(+24.63%)</div>
                            <div class="kpi-value-2 blue">58,611</div>
                            <h5 class="kpi-lable">GST TDS</h5>
                        </div>
                        <div class="kpi-card-box">
                            <div class="kpi-percent red">(+3.38%)</div>
                            <div class="kpi-value-2 blue">64,050</div>
                            <h5 class="kpi-lable">IT TDS/Others</h5>
                        </div>
                    </div>
                    <div class="card-left-line flex">
                        <div class="kpi-card-box">
                            <div class="kpi-percent green">(+6.17%)</div>
                            <div class="kpi-value-2 blue">1,890,288</div>
                            <h5 class="kpi-lable">Total Overdue</h5>
                        </div>
                        <div class="kpi-card-box">
                            <div class="kpi-percent red">(+11.2%)</div>
                            <div class="kpi-value-2 blue">1,536,576</div>
                            <h5 class="kpi-lable">Total Retentions</h5>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- KPI 3 -->
        <div class="kpi-card">
            <h3>Total No. of Projects</h3>
            <div class="kpi-value blue">2,414</div>
            <div class="kpi-change green">(+33)</div>
            <div class="dflex-center">
                <div class="kpi-card-box">
                    <div class="kpi-percent green">(+2%)</div>
                    <div class="kpi-value-2 blue">233</div>
                    <h5 class="kpi-lable">REL</h5>
                </div>
                <div class="kpi-card-box">
                    <div class="kpi-percent green">(+2%)</div>
                    <div class="kpi-value-2 blue">2,096</div>
                    <h5 class="kpi-lable">SO</h5>
                </div>
                <div class="kpi-card-box">
                    <div class="kpi-percent green">(+2%)</div>
                    <div class="kpi-value-2 blue">85</div>
                    <h5 class="kpi-lable">TECO</h5>
                </div>
            </div>
        </div>

        <!-- KPI 4 Wide -->
        <div class="kpi-card wide">
            <div class="dflex">
                <div class="kpi-left-wide">
                    <h3>Sale In Excess</h3>
                    <div class="kpi-value blue">464,507</div>
                    <div class="kpi-change green">(-20.63%)</div>
                </div>
                <div class="kpi-right-half">
                    <div class="kpi-half-box">
                        <h3>Inventory</h3>
                        <div class="kpi-value blue">965,625</div>
                        <div class="kpi-change green">(-10.1%)</div>
                    </div>
                    <div class="card-left-line">
                        <div class="kpi-card-box">
                            <div class="kpi-percent red">(+3.63%)</div>
                            <div class="kpi-value-2 blue">627,228</div>
                            <h5 class="kpi-lable">RM</h5>
                        </div>
                        <div class="kpi-card-box">
                            <div class="kpi-percent red">(+26.3%)</div>
                            <div class="kpi-value-2 blue">242,708</div>
                            <h5 class="kpi-lable">WIP</h5>
                        </div>
                    </div>
                </div>
            </div>

            <div class="dflex">
                <div class="kpi-half-box">
                    <h3>Bill in Excess</h3>
                    <div class="kpi-value blue">1,602,063</div>
                    <div class="kpi-change green">(+9.51%)</div>
                </div>
                <div class="border-box">
                    <h3>NWC Contribution</h3>
                    <div class="kpi-value red">182,284</div>
                </div>
            </div>
        </div>

        <!-- KPI 5 Full Width Repeating Structure -->
        <div class="kpi-card">
    <h3>PR+ Quality %</h3>
    <div class="kpi-row center">
        <div class="kpi-box">
            <div class="kpi-value red">--</div>
            <h5 class="kpi-label">Jan-25</h5>
        </div>
        <div class="kpi-box">
            <div class="kpi-value green">96.15%</div>
            <h5 class="kpi-label">Nov-24</h5>
        </div>
        <div class="kpi-box">
            <div class="kpi-value green">93.13</div>
            <h5 class="kpi-label">YTD</h5>
        </div>
    </div>

    <h3>PR+ Submission %</h3>
    <div class="kpi-row center">
        <div class="kpi-box">
            <div class="kpi-value red">--</div>
            <h5 class="kpi-label">Jan-25</h5>
        </div>
        <div class="kpi-box">
            <div class="kpi-value green">86.15%</div>
            <h5 class="kpi-label">Nov-24</h5>
        </div>
        <div class="kpi-box">
            <div class="kpi-value green">77.13</div>
            <h5 class="kpi-label">YTD</h5>
        </div>
    </div>

    <h3>Approval % (Quarterly)</h3>
    <div class="kpi-row center">
        <div class="kpi-box">
            <div class="kpi-value red">90.2%</div>
            <h5 class="kpi-label">Q1</h5>
        </div>
        <div class="kpi-box">
            <div class="kpi-value green">96.15%</div>
            <h5 class="kpi-label">Q2</h5>
        </div>
        <div class="kpi-box">
            <div class="kpi-value green">93.13</div>
            <h5 class="kpi-label">Q3</h5>
        </div>
        <div class="kpi-box">
            <div class="kpi-value green">83.13</div>
            <h5 class="kpi-label">Q4</h5>
        </div>
    </div>
</div>


        <!-- KPI 6 Wide example repeated -->
        <div class="kpi-card wide">
            <div class="dflex">
                <div class="kpi-half-box">
                    <h3>Actual Revenue Till Dec</h3>
                    <div class="kpi-value green">8,594,344</div>
                    <div class="kpi-change blue">(69.63%)</div>
                </div>
                <div class="kpi-half-box">
                    <h3>Actual Collection Till Dec</h3>
                    <div class="kpi-value green">9,851,948</div>
                    <div class="kpi-change green">(76.63%)</div>
                </div>
            </div>
            <div class="dflex">
                <div class="kpi-half-box">
                    <h3>Actual Revenue (Nov-24)</h3>
                    <div class="kpi-value green">859,319</div>
                    <div class="kpi-change blue">(63.63%)</div>
                </div>
                <div class="kpi-half-box">
                    <h3>Actual Collection (Nov-24)</h3>
                    <div class="kpi-value blue">717,701</div>
                    <div class="kpi-change red">(52.63%)</div>
                </div>
            </div>
            <div class="dflex">
                <div class="kpi-half-box">
                    <h3>Revenue Plan (Jan-25)</h3>
                    <div class="kpi-value blue">547,946</div>
                </div>
                <div class="kpi-half-box">
                    <h3>Collection Plan (Jan-25)</h3>
                    <div class="kpi-value blue">664,507</div>
                </div>
            </div>
            <div class="dflex">
                <div class="kpi-half-box">
                    <h3>Actual Revenue (Jan-25)</h3>
                    <div class="kpi-value green">264,507</div>
                    <div class="kpi-change red">(40.63%)</div>
                </div>
                <div class="kpi-half-box">
                    <h3>Actual Collection (Jan-25)</h3>
                    <div class="kpi-value green">364,507</div>
                    <div class="kpi-change red">(49.63%)</div>
                </div>
            </div>
        </div>

    </div>
`;
const dashboardStyle = ` /* Header */

/* General KPI box styling */
.kpi-row {
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 12px;
    margin-bottom: 15px;
}

.kpi-box {
    flex: 1 1 30%;
    max-width: 30%;
    background: #fff;
    border-radius: 8px;
    padding: 10px;
    box-sizing: border-box;
    text-align: center;
}

.kpi-label { font-size: 12px; margin-top: 4px; color: #666; }
.kpi-value { font-size: 18px; font-weight: bold; margin-bottom: 4px; }

.red { color: #c00000; }
.green { color: #00b050; }
.blue { color: #0047ab; }

/* Special row for quarterly - 4 items in one line */
.quarterly-row .kpi-box {
    flex: 1 1 22%; /* 4 items in one row */
    max-width: 22%;
}

/* Responsive adjustments */
@media (max-width:1024px) {
    .kpi-box { flex: 1 1 45%; max-width: 45%; } /* 2 per row on tablet */
}

@media (max-width:600px) {
    .kpi-box { flex: 1 1 100%; max-width: 100%; } /* 1 per row on mobile */
}


 h1 { color: #0078d4; }
  p { font-size: 16px; }
        .dashboard-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            flex-wrap: wrap;
            margin-bottom: 25px;
        }

        .dashboard-header-left h1 {
            font-size: 26px;
            color: #0078d4;
            margin: 0;
        }

        .dashboard-header-left p {
            font-size: 13px;
            color: #666;
            margin: 2px 0 0;
        }

        .dashboard-header-right {
            text-align: right;
        }

        .dashboard-header-right h2 {
            font-size: 20px;
            color: #444;
            margin: 0;
        }

        .dashboard-header-right p {
            font-size: 13px;
            color: #666;
            margin: 2px 0 0;
        }

        /* KPI Grid */
        .kpi-grid {
            display: flex;
            flex-wrap: wrap;
            gap: 16px;
        }

        .kpi-card {
            background: #fff;
            border-radius: 10px;
            padding: 15px;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.1);
            transition: 0.2s ease-in;
            display: flex;
            flex-direction: column;
            justify-content: center;
            text-align: center;
            flex: 1 1 250px;
            min-width: 250px;
        }

        .kpi-card:hover {
            transform: scale(1.02);
        }

        .kpi-card h3 {
            font-size: 12px;
            text-transform: uppercase;
            margin-bottom: 6px;
            color: #333;
        }

        .kpi-value {
            font-size: 20px;
            font-weight: bold;
            margin-bottom: 4px;
        }

        .kpi-value-2 {
            font-size: 16px;
            font-weight: bold;
            margin-bottom: 4px;
        }

        .kpi-change {
            font-size: 14px;
            margin-bottom: 6px;
        }

        .kpi-lable {
            margin: 0;
        }

        .kpi-percent {
            font-size: 12px;
        }

        .wide {
            flex: 2 1 520px;
        }

        .wide2 {
            flex: 3 1 780px;
        }

        .blue {
            color: #0047ab;
        }

        .green {
            color: #00b050;
        }

        .red {
            color: #c00000;
        }

        .dflex {
            display: flex;
        }

        .dflex-center {
            display: flex;
            justify-content: space-around;
        }

        .split {
            flex: 1;
            display: flex;
            flex-direction: column;
            gap: 10px;
        }
.flex{
display:flex;
}
        .card-left-line {
            border-left: 1px solid #ccc;
            padding-left: 10px;
        }

        .kpi-card-box {
            background: #fff;
            border-radius: 10px;
            padding: 15px;
            margin-bottom: 10px;
            text-align: center;
        }

        /* New classes for former inline styles */
        .kpi-left {
            // width: 25%;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
        }

        .kpi-left-wide {
            width: 40%;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
        }

                     .kpi-right,
                .kpi-left {
                width:100%;
}
        .kpi-right {
            // width: 75%;
            display: flex;
            justify-content: space-between;
            flex-wrap:wrap;
        }

        .kpi-right-half {
            width: 50%;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .kpi-half-box {
            width: 50%;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
        }

        .border-box {
            width: 50%;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            border: 1px solid;
            /* margin: 10px; */
            padding: 40px;
            box-sizing: border-box;
            border-radius: 10px;
        }

        /* Responsive */
        @media (max-width: 1024px) {
            .kpi-card,
            .wide,
            .wide2 {
                flex: 1 1 100%;
            }


            

            .card-left-line {
                border-left: none;
                padding-left: 0;
                display:flex;
            }

            .kpi-left,
            .kpi-left-wide,
            .kpi-right,
            .kpi-right-half,
            .kpi-half-box,
            .border-box {
                width: 100%;
            }
        }`;
const dashboardScript = () => {
    console.log("KPI Dashboard loaded successfully ✅");
};
export default {
    content: dashboardContent,
    style: dashboardStyle,
    script: dashboardScript,
};
//# sourceMappingURL=dashboard.js.map