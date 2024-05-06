import { renderAllCategory, renderTopCategoryView } from './category.js';
import { getUserProfile, getNovel, getListChapterOfNovel, getQueryString } from './const.js';
import { renderTopViewNovel } from './topViewNovel.js';
import { setupSearch } from './search.js';
const access_token = localStorage.getItem("access_token");
const bannerSelector = document.querySelector('.inner-banner');
const tableBodySelector = document.querySelector('.table-body');
const theme = document.querySelector("#theme-link");

const presentDarkMode = localStorage.getItem("dark-mode")

const novelId = getQueryString('id');

const renderNovel = async () => {
    await renderAllCategory();
    await renderTopCategoryView();
    await renderTopViewNovel();


    const novel = (await getNovel(novelId))[0]
    const novelCode = novel.novelCode;
    const listChapter = await getListChapterOfNovel(novelCode);
    let status = ""
    if (novel.status === 0) {
        status = "sắp ra mắt"
    }
    if (novel.status === 1) {
        status = "đang phát hành"
    }
    if (novel.status === 2) {
        status = "đã hoàn thành"
    }
    let listCategory = ""
    novel.category.forEach(category => {
        listCategory += `<a href="#">${category}</a>`
    });

    const bannerHTML = `
    <div class="banner-child">
        <div class="banner-avt">
            <img src="${novel.descriptionImage}" alt="">
            <div class="banner-imp">
                <button class="share">
                    <i class="fa-brands fa-facebook-f"></i>
                    <span>Chia sẻ</span>
                </button>
                <button class="follow">
                    <i class="fa-solid fa-heart"></i>
                    <span>Theo dõi</span>
                </button>
                <button class="unfollow important" style="color: #9c1f66; display: none;">
                    <i class="fa-solid fa-heart"></i>
                    <span>Bỏ Theo dõi</span>
                </button>
            </div>
        </div>
        <div class="banner-desc">
            <h2>${novel.novelName}</h2>
            <div class="in4">
                <ul>
                    <li>Tác giả <a href="#">${novel.authorName}</a></li>
                    <li>Trạng thái <a href="#">${status}</a></li>
                    <li>Lượt xem <a href="#">${novel.view}</a></li>
                    <li>Theo dõi <a href="#">0</a></li>
                    <li>Thể loại ${listCategory}</li>
                </ul>
                <p class="desc">
                    <iframe src="${novel.descriptionURL}" frameborder="0" width="100%" height="100%"></iframe>
                </p>
            </div>
        </div>
    </div>`;
    bannerSelector.innerHTML += bannerHTML;

    listChapter.forEach(chapter => {
        const date = new Date(chapter.updateAt);
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0'); // Thêm số 0 ở đầu nếu tháng chỉ có 1 chữ số
        const day = String(date.getDate()).padStart(2, '0'); // Thêm số 0 ở đầu nếu ngày chỉ có 1 chữ số
        const hours = String(date.getHours()).padStart(2, '0'); // Thêm số 0 ở đầu nếu giờ chỉ có 1 chữ số
        const minutes = String(date.getMinutes()).padStart(2, '0'); // Thêm số 0 ở đầu nếu phút chỉ có 1 chữ số
        const seconds = String(date.getSeconds()).padStart(2, '0'); // Thêm số 0 ở đầu nếu giây chỉ có 1 chữ số
        const formattedDate = `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
        const tableBodyHTML = `
        <tr>
            <td>${chapter.chapterName}</td>
            <td>${formattedDate}</td>
            <td>${chapter.view}</td>
            <td><a href="./read.html?chapterId=${chapter._id}&novelCode=${novelCode}">Đọc</a></td>
        </tr>`;
        tableBodySelector.innerHTML += tableBodyHTML;
    });
}

renderNovel().then(async () => {

    const user = await getUserProfile(access_token)
    let following, check
    if (user._id) {
        following = user.following;
        check = following.includes(novelId);
    }
    // Đặt biến check để kiểm tra xem novelId có trong danh sách following không
    if (check) {
        followButton.style.display = 'none';
        unfollowButton.style.display = 'block';
    }
    const followButton = document.querySelector(".follow");
    const unfollowButton = document.querySelector(".unfollow"); // Define unfollow button here
    followButton.addEventListener("click", async () => {
        if (access_token) {

            followButton.style.display = 'none';
            unfollowButton.style.display = 'block';
            await followNovel(novelId, access_token)
        }
        else {
            alert("Hãy đăng nhập để lưu truyện yêu thích")
        }
    })
    unfollowButton.addEventListener("click", async () => {
        // Toggle the display of unfollow and follow buttons
        unfollowButton.style.display = 'none';
        followButton.style.display = 'block';
        await unfollowNovel(novelId, access_token)
    });
})
    .catch((e) => {
        console.log(e)
    })
document.addEventListener("DOMContentLoaded", setupSearch);

if (presentDarkMode) {
    theme.href = "../assets/css/detail.css";
    // Xử lý khi darkMode là true
} else {
    // Xử lý khi darkMode là false
    theme.href = "../assets/css/detail-dark.css";
}
document.addEventListener('darkModeChange', function (event) {
    const { darkMode } = event.detail;
    // Thực hiện xử lý dựa trên giá trị darkMode
    if (darkMode) {
        theme.href = "../assets/css/detail.css";
        // Xử lý khi darkMode là true
    } else {
        // Xử lý khi darkMode là false
        theme.href = "../assets/css/detail-dark.css";
    }
});