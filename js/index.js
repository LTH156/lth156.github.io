import { renderAllCategory, renderTopCategoryView } from './category.js';
import { fetchData, getDataOrderBy, getQueryString } from './const.js';
import { renderTopViewNovel, slideAction } from './topViewNovel.js';
import { setupSearch } from './search.js';
const listNovelSelector = document.querySelector('.list-novel-new');

console.log(123)
const renderNovel = (data) => {
    // Iterate over each novel in the data
    data.forEach(novel => {
        console.log(novel)
        // Create HTML for each novel
        const novelHTML = `
            <div class="t01"> 
            <div class="item-novel-new col-xl-3 col-md-4 col-6">
                <a href="./detail.html?id=${novel._id}">
                    <div class="image-container">
                        <img src="${novel.descriptionImage}" alt="">
                        <div class="desc-in">
                            <div class="chap item">
                            <i class="fa-solid fa-user-plus"></i>
                                <span>${novel.followed}</span>
                            </div>
                            <div class="view item">
                            <i class="fa-regular fa-eye"></i>
                                <span>${novel.view}</span>
                            </div>
                        </div>
                    </div>
                    <div class="in4">
                        <h2 class="title">${novel.novelName}</h2>
                        <div class="desc">
                            <div class="chap item">
                            <i class="fa-regular fa-user"></i>
                                <span>${novel.authorName}</span>
                            </div>
                            <div class="view item">
                            <i class="fa-solid fa-book"></i>
                                <span>${novel.episodes} tập</span>
                            </div>
                        </div>
                    </div>
                </a>
            </div>
            </div>
        `;
        // Append the novel HTML to the list
        listNovelSelector.innerHTML += novelHTML;
    });
}
const render = async () => {
    const sortBy = getQueryString('sort-by')
    let data
    if (sortBy) {
        data = await getDataOrderBy(sortBy)
    }
    else {
        const listNovelId = getQueryString('listNovelId');
        data = await fetchData(listNovelId);
    }
    renderNovel(data);
    await slideAction();
    await renderTopViewNovel();
    await renderTopCategoryView();
    await renderAllCategory();
}


render().then(() => {

    // console.log('Rendered');
    let thisPage = 1;
    let limit = 16;
    let list = document.querySelectorAll('.t01');
    // console.log(list);

    function loadItem() {
        let beginGet = limit * (thisPage - 1);
        let endGet = limit * thisPage - 1;
        list.forEach((item, key) => {
            if (key >= beginGet && key <= endGet) {
                item.style.display = 'block';
            } else {
                item.style.display = 'none';
            }
        })
        listPage();
    }
    loadItem();
    function listPage() {
        let count = Math.ceil(list.length / limit);
        console.log(count);
        document.querySelector('.listPage').innerHTML = '';

        if (thisPage != 1) {
            let prev = document.createElement('li');
            prev.innerText = '<';
            prev.setAttribute('onclick', "changePage(" + (thisPage - 1) + ")");
            document.querySelector('.listPage').appendChild(prev);
        }
        for (let i = 1; i <= count; i++) {
            let newPage = document.createElement('li');
            newPage.innerText = i;
            if (i == thisPage) {
                newPage.classList.add('active');
            }
            newPage.setAttribute('onclick', "changePage(" + i + ")");
            document.querySelector('.listPage').appendChild(newPage);
        }

        if (thisPage != count) {
            let next = document.createElement('li');
            next.innerText = '>';
            next.setAttribute('onclick', "changePage(" + (thisPage + 1) + ")");
            document.querySelector('.listPage').appendChild(next);

        }
    }
    window.changePage = function (i) {
        thisPage = i;
        console.log(thisPage);
        loadItem();
    }
});
document.addEventListener("DOMContentLoaded", setupSearch);
const theme = document.querySelector("#theme-link");    

const presentDarkMode=localStorage.getItem("dark-mode")
// HEAD
if (presentDarkMode) {
    theme.href = "../assets/css/home.css";
    // Xử lý khi darkMode là true
} else {
    // Xử lý khi darkMode là false
    theme.href = "../assets/css/home-dark.css";
}
// =======

// >>>>>>> c62f64b8349f0f75cdde4427d054d5446f7c9306
document.addEventListener('darkModeChange', function(event) {
    const { darkMode } = event.detail;
    // Thực hiện xử lý dựa trên giá trị darkMode
    if (darkMode) {
        theme.href = "../assets/css/home.css";
        // Xử lý khi darkMode là true
    } else {
        // Xử lý khi darkMode là false
        theme.href = "../assets/css/home-dark.css";
    }
});


