Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        loadTasks();
    }
});

async function loadTasks() {
    const container = document.getElementById("tasks-container");
    if (!container) return;
    container.innerHTML = "<p>Loading your tasks...</p>";

    try {
        // We'll add real Graph code here next
        container.innerHTML = `
        <p><strong>✅ Manifest + pinning is working!</strong></p>
        <p>Next step: Connect to Microsoft Graph To Do and show real tasks sorted by due date.</p>
        <ul style="font-size: 12px; list-style: none; padding-left: 0;">
          <li>☐ Buy groceries — due tomorrow</li>
          <li>☐ Finish project report — due in 2 days</li>
          <li>☐ Call dentist — overdue</li>
        </ul>
        `;
    } catch (error) {
        container.innerHTML = `<p style="color:red;">Error: ${error.message}</p>`;
    }
}
