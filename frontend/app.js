document.getElementById("form").onsubmit = async (e) => {
  e.preventDefault();

  const f = e.target;

  const data = {
    date: f.date.value,
    buyer: f.buyer.value,
    style: f.style.value,
    print_type: f.print_type.value,
    cm_dzn: parseFloat(f.cm_dzn.value),
    quantity: parseInt(f.quantity.value)
  };

  await fetch("/save", {
    method: "POST",
    body: JSON.stringify(data)
  });

  alert("Saved Successfully");
};
