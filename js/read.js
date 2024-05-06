import { renderAllCategory } from './category.js';
import { getChapter, getListChapterOfNovel, getNextChapter, getPreviousChapter, getQueryString } from './const.js';
const contentSelector = document.querySelector('.content');
const previousChapterSelector = document.querySelector('.previous-chapter');
const nextChapterSelector = document.querySelector('.next-chapter');
const listChapterChildSelector = document.querySelector('.list-chapter-1');
const listChapterSelector = document.querySelector('.list-chapter');

let chapterId = getQueryString('chapterId');
let novelCode = getQueryString('novelCode');

const renderContent = async (chapter) => {
    document.title = chapter.chapterName;
    const chapterHTML = `
        <iframe src="${chapter.contentURL}" frameborder="0" width="100%" height="1000px"></iframe>
    `;
    contentSelector.innerHTML = chapterHTML;
}

const renderChapter = async (chapter) => {
    await renderAllCategory();
    await renderContent(chapter);
}
const increaseView = async (chapterId, novelCode) => {
    try {
        const response = await fetch(`https://novel-server-1.onrender.com/novels/increase-view`, {
            method: "PATCH",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                chapterId: chapterId,
                novelCode: novelCode
            })
        });
        const data = await response.json();
        if (!data.success) {
            throw new Error("Failed to increase view count");
        }
    } catch (error) {
        console.error("Error in increaseView function:", error);
        // Handle error appropriately, like throwing it or returning an error object
        throw error;
    }
}

const loadChapter = async (chapterId, novelCode) => {
    await increaseView(chapterId, novelCode);
    const chapter = await getChapter(chapterId, novelCode);
    renderChapter(chapter);
}

previousChapterSelector.addEventListener('click', async () => {
    const previousChapter = await getPreviousChapter(chapterId, novelCode);
    if (previousChapter && previousChapter._id) {
        chapterId = previousChapter._id;
        await loadChapter(chapterId, novelCode);
    } else {
        console.log("No previous chapter available.");
    }
});
nextChapterSelector.addEventListener('click', async () => {
    const nextChapter = await getNextChapter(chapterId, novelCode);
    if (nextChapter && nextChapter._id) {
        chapterId = nextChapter._id;
        await loadChapter(chapterId, novelCode);
    } else {
        console.log("No next chapter available.");
    }
});

loadChapter(chapterId, novelCode).then(async () => {

    const listChapter = await getListChapterOfNovel(novelCode)
    listChapter.forEach((chapter) => {
        const chapterHTML = `
        <li> 
            <a href="./read.html?chapterId=${chapter._id}&novelCode=${chapter.novelCode}">${chapter.chapterName}</a>
        </li>
        `
        listChapterChildSelector.innerHTML += chapterHTML
    })
    listChapterSelector.addEventListener("click", () => {
        // Kiểm tra trạng thái hiện tại của listChapterChildSelector
        if (listChapterChildSelector.style.display === "block") {
            // Nếu đang hiển thị, ẩn đi
            listChapterChildSelector.style.display = "none";
        } else {
            // Nếu đang ẩn, hiển thị lên
            listChapterChildSelector.style.display = "block";
        }
    });
});
const theme = document.querySelector("#theme-link");    

const presentDarkMode=localStorage.getItem("dark-mode")
if (presentDarkMode) {
    theme.href = "../assets/css/read.css";
    // Xử lý khi darkMode là true
} else {
    // Xử lý khi darkMode là false
    theme.href = "../assets/css/read-dark.css";
}
document.addEventListener('darkModeChange', function(event) {
    const { darkMode } = event.detail;
    // Thực hiện xử lý dựa trên giá trị darkMode
    if (darkMode) {
        theme.href = "../assets/css/read.css";
        // Xử lý khi darkMode là true
    } else {
        // Xử lý khi darkMode là false
        theme.href = "../assets/css/read-dark.css";
    }
});