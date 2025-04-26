document.getElementById('upload-form').addEventListener('submit', function() {
    // Show loading message
    document.getElementById('loading').style.display = 'block';
    // Disable the generate button to prevent multiple submissions
    document.getElementById('generate-btn').disabled = true;
});