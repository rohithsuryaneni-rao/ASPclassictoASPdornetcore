// Basic JavaScript functionality
document.addEventListener("DOMContentLoaded", function () {
    const button = document.getElementById("alertButton");
    if (button) {
        button.addEventListener("click", function () {
            alert("Button clicked! This message is triggered by JavaScript.");
        });
    }
});
