const url = "https://novel-server-1.onrender.com"
export const getQueryString = (name) => {
    const search = window.location.search;
    const searchParams = new URLSearchParams(search);
    return searchParams.get(name);
}
export const getListCategory = async () => {
    const response = await fetch(`${url}/novels/get-all-category`, {
        method: "GET"
    });
    return response.json();
}

export const getNovel = async (novelId) => {
    const response = await fetch(`${url}/novels/get-list-novel-by-list-id?listNovelId=${novelId}`, {
        method: "GET"
    });
    return response.json(); // Assuming response is JSON
}

export const getUserProfile = async (access_token) => {
    const user = await fetch(`${url}/users/get-me`, {
        method: "GET",
        headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${access_token}`
        }
    });
    return user.json()
}
export const followNovel = async (novelId, access_token) => {
    const result = await fetch(`${url}/users/follow-novel`, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${access_token}`
        },
        body: JSON.stringify({
            novelId
        })
    });
    return result.json()
}
export const unfollowNovel = async (novelId, access_token) => {
    const result = await fetch(`${url}/users/unfollow-novel`, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${access_token}`
        },
        body: JSON.stringify({
            novelId
        })
    });
    return result.json()
}

export const fetchData = async (listId = "") => {
    let queryRequest = "";
    if (listId) {
        queryRequest = listId.split(',').map((id) => `${id}`).join("&listNovelId=");
    }
    const response = await fetch(`${url}/novels/get-list-novel-by-list-id?listNovelId=${queryRequest}`, {
        method: "GET"
    });
    return response.json();
};
export const getDataOrderBy = async (sortBy) => {
    let queryRequest = "";
    if (sortBy === "view") {
        queryRequest = "get-novel-sorted-by-view"
    }
    else if (sortBy === "alphabet") {
        queryRequest = "get-novel-sorted-alphabetically"

    }
    const response = await fetch(`${url}/novels/${queryRequest}`, {
        method: "GET"
    });
    return response.json();
}
export const getQueryParameter = (name) => {
    const searchParams = new URLSearchParams(window.location.search);
    return searchParams.get(name);
}


export const getListChapterOfNovel = async (novelCode) => {
    const response = await fetch(`${url}/novels/get-chapter-in-novel?novelCode=${novelCode}`, {
        method: "GET"
    });
    return response.json();
}
export const getChapter = async (chapterId, novelCode) => {
    const response = await fetch(`${url}/novels/get-chapter?chapterId=${chapterId}&novelCode=${novelCode}`, {
        method: "GET"
    });
    return response.json();
}

export const getPreviousChapter = async (chapterId, novelCode) => {
    const response = await fetch(`${url}/novels/get-previous-chapter?chapterId=${chapterId}&novelCode=${novelCode}`, {
        method: "GET"
    });
    return response.json();
}

export const getNextChapter = async (chapterId, novelCode) => {
    const response = await fetch(`${url}/novels/get-next-chapter?chapterId=${chapterId}&novelCode=${novelCode}`, {
        method: "GET"
    });
    return response.json();
}
export const searchNovel = async (description) => {
    const response = await fetch(`${url}/novels/search-novel`, {
        method: "POST",
        headers: {
            "Content-Type": "application/json"
        },
        body: JSON.stringify({
            description
        })
    });
    return response.json();
}
