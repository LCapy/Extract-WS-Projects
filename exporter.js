document.getElementById("upload-form").addEventListener("submit", async function (e) {
    e.preventDefault();

    const token = document.getElementById("token-input").value.trim();
    const cookiesFile = document.getElementById("cookies-file").files[0];

    if (!token || !cookiesFile) {
        alert("Please enter the token and upload the cookies file.");
        return;
    }

    const reader = new FileReader();
    reader.onload = async function () {
        const cookies = JSON.parse(reader.result);
        const headers = { 'User-Agent': 'Mozilla/5.0' };
        const url = `https://ws.corp.adobe.com/ws-legacy/assignments_projects?&token=${token}&navigationPanel=true`;

        const response = await fetch(url, {
            method: 'GET',
            headers: headers,
            credentials: 'include'
        });

        if (response.ok) {
            const htmlContent = await response.text();
            extractAndDownloadData(htmlContent);
        } else {
            document.getElementById("status-message").textContent = "Failed to retrieve data.";
        }
    };
    reader.readAsText(cookiesFile);
});

function extractAndDownloadData(htmlContent) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(htmlContent, "text/html");
    const projectsData = [];

    const rows = doc.querySelectorAll('tbody[id^="defaultPagedTableId"] > tr.table_bg_color');

    let currentProjectName, currentCreatedDate, currentSourceLanguage;

    rows.forEach((row) => {
        const columns = row.querySelectorAll("td.paged_table_content");
        if (columns.length === 1) {
            const text = columns[0].textContent.trim();
            if (text.includes("created")) {
                currentProjectName = text.split(" (created")[0].trim();
                currentCreatedDate = text.split("created ")[1].split(",")[0].trim();
                currentSourceLanguage = text.split("source:")[1].split(")")[0].trim();
            }
        } else if (columns.length > 1) {
            const targetLanguage = columns[0].querySelector("a:last-child")?.textContent.trim() || "Unknown";
            const totalWords = columns[2]?.textContent.trim() || "N/A";
            const stepStatus = columns[3]?.textContent.includes("Exported") ? "Exported" : "In Progress";
            const claimedBy = columns[5]?.textContent.trim() || "Unassigned";
            const phase = columns[9]?.textContent.trim() || "N/A";
            const step = columns[10]?.textContent.trim() || "N/A";

            projectsData.push({
                "Project Name": currentProjectName,
                "Created": currentCreatedDate,
                "Source Language": currentSourceLanguage,
                "Target Language": targetLanguage,
                "Total Words": totalWords,
                "Step Status": stepStatus,
                "Claimed By": claimedBy,
                "Phase": phase,
                "Step": step
            });
        }
    });

    generateExcelFile(projectsData);
}

function generateExcelFile(data) {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, worksheet, "WorldServer Projects");
    XLSX.writeFile(workbook, "projects_report.xlsx");
}
