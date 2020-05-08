function downloadMyExcel() {
    var workbook = new ExcelJS.Workbook();
    var filename = "AccessBotPluginReport.xlsx";

    //worksheets
    var summaryWorksheet = workbook.addWorksheet("Summary");
    var detailsWorksheet = workbook.addWorksheet("Details");

    //details worksheet data
    var detailsHeader = [
        "Severity",
        "Rule Name",
        "Issue Description",
        "Remediation Suggestion",
        "Source",
        "Element",
        "Tags"
    ];
    var detailsData = [
        {
            "Serverity": "minor",
            "Rule Name": "aria-allowed-role",
            "Issue Description":
                "Ensures role attribute has an appropriate value for the element",
            "Remediation Suggestion":
                "Fix any of the following:\n  ARIA role button is not allowed for given element",
            "Source": "#toctogglecheckbox",
            "Element":
                "<input type=\"checkbox\" role=\"button\" id=\"toctogglecheckbox\" class=\"toctogglecheckbox\" style=\"display:none\">",
            "Tags": "cat.aria, best-practice"
        },
        {
            "Serverity": "minor",
            "Rule Name": "aria-allowed-role",
            "Issue Description":
                "Ensures role attribute has an appropriate value for the element",
            "Remediation Suggestion":
                "Fix any of the following:\n  ARIA role button is not allowed for given element",
            "Source": "#toctogglecheckbox",
            "Element":
                "<input type=\"checkbox\" role=\"button\" id=\"toctogglecheckbox\" class=\"toctogglecheckbox\" style=\"display:none\">",
            "Tags": "cat.aria, best-practice"
        },
        {
            "Serverity": "minor",
            "Rule Name": "aria-allowed-role",
            "Issue Description":
                "Ensures role attribute has an appropriate value for the element",
            "Remediation Suggestion":
                "Fix any of the following:\n  ARIA role button is not allowed for given element",
            "Source": "#toctogglecheckbox",
            "Element":
                "<input type=\"checkbox\" role=\"button\" id=\"toctogglecheckbox\" class=\"toctogglecheckbox\" style=\"display:none\">",
            "Tags": "cat.aria, best-practice"
        }
    ];

    //summary worksheet data
    var summary = {
        "Scanned URL": "https://en.wikipedia.org/wiki/Late_Cretaceous",
        "User Agent":
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4128.3 Safari/537.36",
        "Orientation Type": "landscape-primary",
        "Orientation Angle": 0,
        "Window Width": 1536,
        "Window Height": 576,
        "Total Issues": 9,
        "Total Violations": 5,
        Inapplicable: 26,
        Passed: 44,
    };
    var violationsBySeverity = {Minor: 13, Major: 8, Critical: 1};
    var violationsByRuleCount = {
        "aria-allowed-role": 1,
        label: 1,
        "link-name": 2,
        list: 1,
        region: 2,
        "color-contrast": 3,
        "identical-links-same-purpose": 12,
    };

    //summary worksheet data filling
    summaryWorksheet.addRow(["Scan Summary:"]).font = {
        name: 'Calibri',
        size: 16,
        bold: true
    };
    summaryWorksheet.addRow()
    for (let key in summary) {
        summaryWorksheet.addRow([key, summary[key]])
    }

    summaryWorksheet.addRow()
    summaryWorksheet.addRow()
    summaryWorksheet.addRow(["Violations by rule:"]).font = {
        name: 'Calibri',
        size: 16,
        bold: true
    };

    for (let key in violationsByRuleCount) {
        summaryWorksheet.addRow([key, violationsByRuleCount[key]])
    }

    summaryWorksheet.addRow()
    summaryWorksheet.addRow()
    summaryWorksheet.addRow(["Violations By Severity:"]).font = {
        name: 'Calibri',
        size: 16,
        bold: true
    };

    for (let key in violationsBySeverity) {
        summaryWorksheet.addRow([key, violationsBySeverity[key]])
    }

    //summary worksheet columns alignment
    summaryWorksheet.getColumn(1).width = 50;
    summaryWorksheet.getColumn(2).width = 100;
    summaryWorksheet.getColumn(1).alignment = {horizontal: 'left', wrapText: true};
    summaryWorksheet.getColumn(2).alignment = {horizontal: 'right', wrapText: true};

    //Add Header Row
    detailsWorksheet.addRow(detailsHeader);

    // Add Data and Formatting
    for (var i = 0; i < detailsData.length; i++) {
        obj = [detailsData[i].Serverity, detailsData[i]["Rule Name"], detailsData[i]["Issue Description"], detailsData[i]["Remediation Suggestion"],
            detailsData[i].Source, detailsData[i].Element, detailsData[i].Tags]
        row = detailsWorksheet.addRow(obj);
        row.height = 180
    }

    detailsWorksheet.getColumn(1).alignment = {vertical: 'bottom', horizontal: 'left', wrapText: true};
    detailsWorksheet.getColumn(2).alignment = {vertical: 'bottom', horizontal: 'left', wrapText: true};
    detailsWorksheet.getColumn(3).alignment = {vertical: 'top', horizontal: 'left', wrapText: true};
    detailsWorksheet.getColumn(4).alignment = {vertical: 'bottom', horizontal: 'left', wrapText: true};
    detailsWorksheet.getColumn(5).alignment = {vertical: 'bottom', horizontal: 'left', wrapText: true};
    detailsWorksheet.getColumn(6).alignment = {vertical: 'top', horizontal: 'left', wrapText: true};
    detailsWorksheet.getColumn(7).alignment = {vertical: 'top', horizontal: 'left', wrapText: true};

    detailsWorksheet.getColumn(1).width = 25;
    detailsWorksheet.getColumn(2).width = 30;
    detailsWorksheet.getColumn(3).width = 50;
    detailsWorksheet.getColumn(4).width = 50;
    detailsWorksheet.getColumn(5).width = 50;
    detailsWorksheet.getColumn(6).width = 50;
    detailsWorksheet.getColumn(7).width = 50;

    detailsWorksheet.getColumn(1).font = {
        name: 'Calibri',
        size: 11,
    };
    detailsWorksheet.getColumn(2).font = {
        name: 'Calibri',
        size: 11,
    };
    detailsWorksheet.getColumn(3).font = {
        name: 'Calibri',
        size: 11,
    };
    detailsWorksheet.getColumn(4).font = {
        name: 'Calibri',
        size: 11,
    };
    detailsWorksheet.getColumn(5).font = {
        name: 'Calibri',
        size: 11,
    };
    detailsWorksheet.getColumn(6).font = {
        name: 'Calibri',
        size: 11,
    };
    detailsWorksheet.getColumn(7).font = {
        name: 'Calibri',
        size: 11,
    };
    //Generate Excel File with given name
    workbook.xlsx
        .writeBuffer({
            base64: true,
        })
        .then(function (data) {
            // build anchor tag and attach file (works in chrome)
            var a = document.createElement("a");
            var data = new Blob([data], {
                type:
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            });

            var url = URL.createObjectURL(data);
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            setTimeout(function () {
                document.body.removeChild(a);
                window.URL.revokeObjectURL(url);
            }, 0);
        })
        .catch(function (error) {
            console.log(error.message);
        });
}
