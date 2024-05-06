// Selectors
const listCategorySelector = document.querySelector('#category-list');
const listCategoryOnHeader = document.querySelector('.list');
const maxCategoriesToShow = 8; // Maximum number of categories to show
import { getListCategory } from "./const.js";

// Fetch categories data
const renderTopCategoryView = async () => {
    const data = await getListCategory();
    // Extract first maxCategoriesToShow categories
    const categoriesToShow = data.slice(0, maxCategoriesToShow);
    // Render categories in the category list
    categoriesToShow.forEach(category => {
        const queryRequest = category.novelId.map((id) => `${id}`).join(",");
        const categoryHTML = `
        <a href="./index.html?listNovelId=${queryRequest}">
            <div class="category-item" style="background-image: url(${category.imgUrl})">
                <span>${category.categoryName}</span>
            </div>
        </a>
        `;
        listCategorySelector.innerHTML += categoryHTML;
    });
}
const renderAllCategory = async () => {
    const data = await getListCategory();
    // Render categories in the header
    data.forEach(category => {
        const queryRequest = category.novelId.map((id) => `${id}`).join(",");
        const categoryHTML = `
        <li>
            <a href="./index.html?listNovelId=${queryRequest}">
                <div class="category-item" style="background-image: url(${category.imgURL})">
                <span>${category.categoryName}</span>
                </div>
            </a>
        </li>
    
    
        `;
        listCategoryOnHeader.innerHTML += categoryHTML;
    });
}
export { renderAllCategory, renderTopCategoryView };

