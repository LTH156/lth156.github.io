import { searchNovel } from './const.js';

export function setupSearch() {
    const searchInput = document.getElementById("searchInput");
    const searchButton = document.getElementById("searchButton");

    async function searchHandler() {
        const description = searchInput.value;
        
        try {
            const result = await searchNovel(description);
            
            // Lấy danh sách _id từ kết quả tìm kiếm
            const listId = result.map(novel => novel._id);
            
            // Tạo query string từ danh sách _id
            const queryRequest = listId.join(",");
            
            // Chuyển hướng người dùng đến index.html với query string
            window.location.href = `./index.html?listNovelId=${queryRequest}`;
        } catch (error) {
            console.error("Đã xảy ra lỗi khi tìm kiếm:", error);
        }
    }

    searchButton.addEventListener("click", searchHandler);

    searchInput.addEventListener("keyup", function(event) {
        if (event.keyCode === 13) {
            searchHandler();
        }
    });
}
