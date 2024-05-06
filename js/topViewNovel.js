const topViewNovelSelector = document.querySelector('.view-list');
const slideSelector = document.querySelector('.slider');
import { fetchData } from "./const.js";
const renderTopViewNovel = async () => {
    const novels = await fetchData();
    let count = 0; // Biến đếm số lượng truyện đã thêm
    novels.forEach(novel => {
        if (count < 5) { // Kiểm tra xem đã đủ 10 truyện chưa
            const novelHTML = `
            <div class="view-item">
            <a href="#" class="name">${novel.novelName}</a>
            <div class="main-novel">
                <div class="sub-novel">
                    <div class="img">
                        <img src="${novel.descriptionImage}" alt="">
                    </div>
                    <div class="content-novel">
                        <div class="ct-in4">
                            <div class="view item">
                            <i class="fa-regular fa-eye"></i>
                                <span>${novel.view} lượt xem</span>
                            </div>
                            <div class="chap item">
                            <i class="fa-solid fa-user-plus"></i>
                                <span>${novel.followed} theo dõi</span>
                            </div>
                            <a class="xem_ngay item" href = './detail.html?id=${novel._id}'>XEM NGAY</a>
                        </div>
                    </div>
                </div>
            </div>
        </div>
            `;
            topViewNovelSelector.innerHTML += novelHTML;



            count++; // Tăng biến đếm sau khi thêm một truyện
        }
    });
}

const renderSlide = async () => {
    const novels = await fetchData();
    let count = 0; // Biến đếm số lượng truyện đã thêm
    novels.forEach(novel => {
        if (count < 10) { // Kiểm tra xem đã đủ 5 truyện chưa
            const slideHTML = `
            <a href="./detail.html?id=${novel._id}">
            <div class="slide slide-1">
                <div class="image-container">
                    <img src="${novel.descriptionImage}"
                        alt="Image 1">
                </div>
                <div class="content">
                    <h4 class="title">${novel.novelName}</h4>
                    <div class="in4">
                        <div class="chap item">
                        <i class="fa-solid fa-book"></i>
                            <span>${novel.episodes} tập</span>
                        </div>
                        <div class="view item">
                        <i class="fa-regular fa-eye"></i>
                            <span>${novel.view}</span>
                        </div>
                        <div class="calendar item">
                        
                        <i class="fa-regular fa-user"></i>
                            <span>${novel.authorName}</span>
                        </div>
                        <div class="follow item">
                        <i class="fa-solid fa-user-plus"></i>
                            <span>${novel.followed}</span>
                        </div>
                    </div>
                </div>
            </div>
            </a>

            
            `;
            slideSelector.innerHTML += slideHTML;
        }
    });
}
const slideAction = async () => {
    await renderSlide();
    const sliderContainer = document.querySelector('.slider-container');
    const slider = document.querySelector('.slider');
    const slides = document.querySelectorAll('.slide');

    let currentSlide = 0;
    let slideCount = 0;

    function nextSlide() {
        const slideWidth = slides[0].offsetWidth + 20; // 20px là khoảng cách giữa các ảnh
        const containerWidth = sliderContainer.offsetWidth;
        const maxOffset = (slides.length - 1) * slideWidth;

        if (maxOffset > containerWidth) {
            currentSlide = (currentSlide + 1) % slides.length;
            updateSlider();
            slideCount++;

            if (slideCount === 5) {
                slideCount = 0;
                currentSlide = 0;
                slider.style.transform = `translateX(0)`;
            }
        } else {
            currentSlide = 0;
            slider.style.transform = `translateX(0)`;
        }
    }

    function updateSlider() {
        const offset = -currentSlide * (slides[0].offsetWidth + 20); // 20px là khoảng cách giữa các ảnh
        slider.style.transform = `translateX(${offset}px)`;
    }
    // Tự động chuyển slide sau mỗi 1 giây
    setInterval(nextSlide, 3000);
}
export { renderTopViewNovel, slideAction };

