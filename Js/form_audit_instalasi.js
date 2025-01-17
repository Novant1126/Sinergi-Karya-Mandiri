document.addEventListener("DOMContentLoaded", function() {
    const openCameraButton = document.getElementById("openCameraInstalasi");
    const captureImageButton = document.getElementById("captureImageInstalasi");
    const saveImageButton = document.getElementById("saveImageInstalasi");
    const videoElement = document.getElementById("videoInstalasi");
    const canvasElement = document.getElementById("canvasInstalasi");
    const capturedImageElement = document.getElementById("capturedImageInstalasi");
    let stream;

    // Fungsi untuk membuka kamera
    function openCamera(event) {
        event.preventDefault(); // Mencegah form untuk submit saat tombol diklik
        navigator.mediaDevices.getUserMedia({ video: true })
            .then(function(mediaStream) {
                stream = mediaStream;
                videoElement.srcObject = mediaStream;
                videoElement.play();
                document.getElementById("captureImageInstalasi").classList.remove("d-none");
                document.getElementById("saveImageInstalasi").classList.remove("d-none");
            })
            .catch(function(error) {
                console.error("Error accessing camera: ", error);
            });
    }

    // Fungsi untuk menangkap gambar dari video
    function captureImage(event) {
        event.preventDefault(); // Mencegah form untuk submit saat tombol diklik
        const context = canvasElement.getContext("2d");
        context.drawImage(videoElement, 0, 0, canvasElement.width, canvasElement.height);
        capturedImageElement.src = canvasElement.toDataURL("image/png");
        capturedImageElement.classList.remove("d-none");
        canvasElement.classList.add("d-none");
        videoElement.classList.add("d-none");
    }

    // Fungsi untuk menyimpan gambar
    function saveImage(event) {
        event.preventDefault(); // Mencegah form untuk submit saat tombol diklik
        const dataUrl = capturedImageElement.src;
        // Di sini, Anda dapat mengirim data gambar ke server atau menyimpannya sesuai kebutuhan
        console.log("Image captured: ", dataUrl);
    }

    openCameraButton.addEventListener("click", openCamera);
    captureImageButton.addEventListener("click", captureImage);
    saveImageButton.addEventListener("click", saveImage);
});
