import { renderAllCategory, renderTopCategoryView } from './category.js';
import { getUserProfile,fetchData } from './const.js';
import { renderTopViewNovel } from './topViewNovel.js';
const listNovelSelector = document.querySelector('.list-novel-new');
const access_token = localStorage.getItem("access_token");


const renderNovel = (data) => {
    // Iterate over each novel in the data
    data.forEach(novel => {
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
                            <i class="fa-solid fa-user-plus"></i>
                                <span>${novel.authorName}</span>
                            </div>
                            <div class="view item">
                            <i class="fa-solid fa-book"></i>
                                <span>${novel.episodes}Tập</span>
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
    const user=await getUserProfile(access_token)
    const {following}=user
    const data=await fetchData(following)
    renderNovel(data);
    await renderTopViewNovel();
    await renderTopCategoryView();
    await renderAllCategory();
}
render().then(() => {
 
    let thisPage = 1;
    let limit = 16;
    let list = document.querySelectorAll('.t01');

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
        loadItem();
    }
});


