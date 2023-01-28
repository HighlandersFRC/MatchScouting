function setUpGoogleSheets() {
    const scriptURL = 'https://script.google.com/macros/s/AKfycbzoy0_y5foON2jNUlNAoEwSSWqaY3GZY4LiGX8hz7uX7_hozd73ISuSc4J-scfakHGS/exec'
    const form = document.querySelector('#scoutingForm')
    const btn = document.querySelector('#submit')
    alert("Hello")
    return true;


    
    form.addEventListener('submit', e => {
      e.preventDefault()
      btn.disabled = true
      btn.innerHTML = "Sending..."

      let fd = getData(false)
      for (const [key, value] of fd) {
        console.log(`${key}: ${value}\n`);
      }

      fetch(scriptURL, { method: 'POST', mode: 'no-cors', body: fd })
        .then(response => { 
              alert('Success!', response) })
        .catch(error => {
              alert('Error!', error.message)})

      btn.disabled = false
      btn.innerHTML = "Send to Google Sheets"
    })
}
