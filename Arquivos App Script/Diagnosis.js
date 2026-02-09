
function handleDiagnostic() {
    try {
        const props = PropertiesService.getScriptProperties().getProperties();
        const dbUrlProd = props['DB_URL_PROD'] || "N/A";
        const dbUrlStaging = props['DB_URL_STAGING'] || "N/A";
        const envVar = props['ENV'] || "N/A";

        const ss = getDatabaseConnection();
        const sheet = ss.getSheetByName("Solicitações");

        let lastRow = 0;
        let lastId = "N/A";
        let headers = [];

        if (sheet) {
            lastRow = sheet.getLastRow();
            if (lastRow > 0) {
                headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
                if (lastRow > 1) {
                    lastId = sheet.getRange(lastRow, 1).getValue();
                }
            }
        }

        return {
            status: "OK",
            environment: {
                detected_prod: isProduction(),
                env_var: envVar
            },
            config: {
                db_url_prod_set: dbUrlProd !== "N/A",
                db_url_staging_set: dbUrlStaging !== "N/A"
            },
            connected_sheet: {
                name: ss.getName(),
                id: ss.getId(),
                has_queue_sheet: !!sheet,
                total_rows: lastRow,
                last_request_id: lastId,
                headers_sample: headers.slice(0, 3)
            }
        };
    } catch (e) {
        return {
            status: "ERROR",
            message: e.message,
            stack: e.stack
        };
    }
}
