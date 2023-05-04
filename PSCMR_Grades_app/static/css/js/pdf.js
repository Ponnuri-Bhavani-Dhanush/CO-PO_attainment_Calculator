window.onload = function(){
    document.getElementById("download")
    .addEventListener("click",()=>{
        const invoice = this.document.getElementById("main-div");
        var opt = {
            filename: 'myfile.pdf',
            image: { type: 'pdf', quality: 1 },
            html2canvas: { scale: 1 },
            jsPDF: { unit: 'in', format: 'A4', orientation: 'p' }
        };
        html2pdf().from(invoice).set(opt).save();
        // html2pdf().from(invoice).save();
    })
}