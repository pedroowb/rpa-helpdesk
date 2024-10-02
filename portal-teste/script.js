document.getElementById("submit-btn").addEventListener("click", function(event) {
    event.preventDefault();
    var form = document.getElementById("myForm");
    form.reset();
  });