/*
    employmentRadios.js

    Purpose:
    - Toggle visibility of form sections based on radio inputs on the customer form.
    - Expected to run after the DOM is loaded and requires inputs with name="employmentType" and name="form16".
    - Controls elements with IDs: `salariedSection`, `selfEmpSection`, `form16Yes`, `form16No`.
*/

// Employment type toggle
document.querySelectorAll('input[name="employmentType"]').forEach(radio => {
    radio.addEventListener("change", function () {
        document.getElementById("salariedSection").style.display =
            this.value === "salaried" ? "block" : "none";
        document.getElementById("selfEmpSection").style.display =
            this.value === "selfEmployed" ? "block" : "none";
    });
});

// Form 16 toggle
document.querySelectorAll('input[name="form16"]').forEach(radio => {
    radio.addEventListener("change", function () {
        document.getElementById("form16Yes").style.display =
            this.value === "yes" ? "block" : "none";
        document.getElementById("form16No").style.display =
            this.value === "no" ? "block" : "none";
    });
});