Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        loadTasks();
    }
});

async function loadTasks() {
    const container = document.getElementById("tasks-container");
    container.innerHTML = "<p>Loading your tasks...</p>";

    try {
        // We'll add real Graph code here next
        container.innerHTML = `
        <p><strong>✅ Manifest + pinning is working!</strong></p>
        <p>Next step: Connect to Microsoft Graph To Do and show real tasks sorted by due date.</p>
        `;
    } catch (error) {
        container.innerHTML = `<p style="color:red;">Error: ${error.message}</p>`;
    }
}
