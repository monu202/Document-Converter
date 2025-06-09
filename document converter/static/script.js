document.addEventListener('DOMContentLoaded', () => {
    const hero = document.getElementById('hero');
    const getStartedButton = document.getElementById('getStartedButton');
    let isHeroShrunk = false;

    // Handle Get Started button click
    getStartedButton.addEventListener('click', () => {
        if (isHeroShrunk) {
            hero.classList.remove('shrink');
            window.scrollTo({ top: 0, behavior: 'smooth' });
        } else {
            hero.classList.add('shrink');
            document.getElementById('converterForm').scrollIntoView({ behavior: 'smooth' });
        }
        isHeroShrunk = !isHeroShrunk;
    });

    // Handle navigation links dynamically
    document.querySelectorAll('.navbar a').forEach(link => {
        link.addEventListener('click', (event) => {
            event.preventDefault(); // Prevent default anchor behavior
            const targetId = link.getAttribute('data-target'); // Get the target section ID
            
            // Hide all sections
            document.querySelectorAll('.info-section').forEach(section => {
                section.classList.add('hidden'); // Add the "hidden" class to hide sections
            });

            // Show the targeted section
            const targetSection = document.getElementById(targetId);
            if (targetSection) {
                targetSection.classList.remove('hidden'); // Remove the "hidden" class to display the section
            }
        });
    });
});
