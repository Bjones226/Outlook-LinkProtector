document.addEventListener("DOMContentLoaded", function () {
  const linksContainer = document.getElementById("links");
  const unlockAllLinksButton = document.getElementById("unlockAllLinks");

  Office.onReady(() => {
    const item = Office.context.mailbox.item;
    item.body.getAsync(Office.CoercionType.Text, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const bodyText = result.value;
        const urls = bodyText.match(/\bhttps?:\/\/[^\s<>"']+/gi) || [];

        urls.forEach((url) => {
          const row = document.createElement("div");
          row.className = "link-row";

          const lockIcon = document.createElement("span");
          lockIcon.textContent = "🔒";

          const label = document.createElement("span");
          label.textContent = url;

          const unlockBtn = document.createElement("button");
          unlockBtn.textContent = "Unlock";
          unlockBtn.addEventListener("click", () => {
            row.innerHTML = "✅ ";
            const link = document.createElement("a");
            link.href = url;
            link.textContent = url;
            link.target = "_blank";
            row.appendChild(link);
          });

          row.appendChild(lockIcon);
          row.appendChild(label);
          row.appendChild(unlockBtn);
          linksContainer.appendChild(row);
        });
      }
    });
  });

  unlockAllLinksButton.addEventListener("click", () => {
    linksContainer.querySelectorAll("button").forEach((btn) => btn.click());
  });
});
