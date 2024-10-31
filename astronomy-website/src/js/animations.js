// This file contains JavaScript code for interactive animations to enhance user experience.

document.addEventListener('DOMContentLoaded', function() {
    const elements = document.querySelectorAll('.animate');

    elements.forEach(element => {
        element.addEventListener('mouseover', () => {
            element.classList.add('hover-animation');
        });

        element.addEventListener('mouseout', () => {
            element.classList.remove('hover-animation');
        });
    });

    const scrollElements = document.querySelectorAll('.scroll-animation');

    const elementInView = (el, dividend = 1) => {
        const elementTop = el.getBoundingClientRect().top;
        return (
            elementTop <= (window.innerHeight || document.documentElement.clientHeight) / dividend
        );
    };

    const displayScrollElement = (element) => {
        element.classList.add('visible');
    };

    const handleScrollAnimation = () => {
        scrollElements.forEach((el) => {
            if (elementInView(el, 1.25)) {
                displayScrollElement(el);
            }
        });
    };

    window.addEventListener('scroll', () => {
        handleScrollAnimation();
    });
});