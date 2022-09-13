let intro = document.querySelector('.intro');
let logo = document.querySelector('.logo-header');
let logoSpan = document.querySelectorAll('.logo');

window.addEventListener('DOMContentLoaded', ()=>{
    setTimeout(()=>{
        logoSpan.forEach((span, i, )=>{
            setTimeout(()=>{
                span.classList.add('active');
            }, (i + 1) * 400)
        })

        setTimeout(()=>{
            logoSpan.forEach((span, i, )=>{
                setTimeout(()=>{
                    span.classList.remove('active');
                    span.classList.add('fade');
                }, (i + 1) * 50)
            })
        }, 2000);

        setTimeout(()=>{
            intro.style.top = '-100vh';
        }, 2300);
    })
})