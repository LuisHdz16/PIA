document.addEventListener("DOMContentLoaded", function () {
    let currentSlide = 0;
    const slides = document.querySelectorAll(".slide");
    const totalSlides = slides.length;

    function showSlide(index) {
        slides.forEach((slide) => {
            slide.style.display = "none";
        });

        slides[index].style.display = "block";
    }

    function nextSlide() {
        currentSlide = (currentSlide + 1) % totalSlides;
        showSlide(currentSlide);
    }

    function prevSlide() {
        currentSlide = (currentSlide - 1 + totalSlides) % totalSlides;
        showSlide(currentSlide);
    }

    function autoSlide() {
        nextSlide();
    }

    // Mostrar la primera imagen al cargar la página
    showSlide(currentSlide);

    // Configurar temporizador para cambio automático cada 3 segundos (3000 milisegundos)
    const intervalId = setInterval(autoSlide, 3000);

    // Agregar event listeners para los botones de siguiente y anterior
    document.getElementById("nextBtn").addEventListener("click", function () {
        clearInterval(intervalId); // Detener temporizador al hacer clic manual
        nextSlide();
    });

    document.getElementById("prevBtn").addEventListener("click", function () {
        clearInterval(intervalId); // Detener temporizador al hacer clic manual
        prevSlide();
    });
});
