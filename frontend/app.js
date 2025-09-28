document.getElementById('selectFileBtn').onclick = async function() {
    if (window.pywebview) {
        const result = await window.pywebview.api.read_data('');
        document.getElementById('infoArea').textContent = result;
    } else {
        alert('pywebview API no disponible');
    }
};

document.getElementById('exportBtn').onclick = async function() {
    if (window.pywebview) {
        const result = await window.pywebview.api.export_excel();
        document.getElementById('infoArea').textContent += '\n' + result;
    } else {
        alert('pywebview API no disponible');
    }
};
