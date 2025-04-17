let fetchedData = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("fetchData").onclick = fetchData;
        document.getElementById("displayData").onclick = displayData;
        document.getElementById("insertHtml").onclick = insertHtml;
    }
});

async function fetchData() {
    const token = document.getElementById("token").value;
    const url = document.getElementById("url").value;
    const status = document.getElementById("status");

    if (!token || !url) {
        status.innerText = "Please enter both token and URL.";
        return;
    }

    try {
        const response = await fetch(url, {
            headers: {
                Authorization: `Bearer ${token}`
            }
        });
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        fetchedData = await response.text();
        status.innerText = "Data fetched successfully!";
    } catch (error) {
        status.innerText = `Error fetching data: ${error.message}`;
        fetchedData = null;
    }
}

function displayData() {
    const status = document.getElementById("status");
    if (!fetchedData) {
        status.innerText = "Please fetch data first.";
        return;
    }
    status.innerText = `Fetched Data: ${fetchedData}`;
}

async function insertHtml() {
    const status = document.getElementById("status");
    if (!fetchedData) {
        status.innerText = "Please fetch data first.";
        return;
    }

    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            body.insertHtml(fetchedData, Word.InsertLocation.end);
            await context.sync();
            status.innerText = "HTML content inserted into document.";
        });
    } catch (error) {
        status.innerText = `Error inserting HTML: ${error.message}`;
    }
}