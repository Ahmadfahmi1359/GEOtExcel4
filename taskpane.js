Office.onReady(() => {
  document.querySelectorAll(".btn").forEach(button => {
    button.addEventListener("click", () => {
      const value = parseFloat(button.dataset.value);
      adjustValue(value);
    });
  });
});

async function adjustValue(amount) {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("values");
    await context.sync();
    const currentValue = parseFloat(range.values[0][0]);
    if (!isNaN(currentValue)) {
      range.values = [[currentValue + amount]];
    }
    await context.sync();
  });
}
