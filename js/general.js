const darkModeSelector = document.querySelector(".dark-mode")
const theme = document.querySelector("#theme-link");
const access_token = localStorage.getItem("access_token")
const loginSelector = document.querySelector(".login")
const logoutSelector = document.querySelector(".logout-after")
const helloUserSelector = document.querySelector(".hello_user")

import { getUserProfile } from "./const.js";


darkModeSelector.addEventListener("click", function () {
    // Lấy giá trị darkMode từ localStorage và chuyển đổi thành boolean
    const darkMode = localStorage.getItem("dark-mode")

    // Cập nhật giá trị darkMode trong localStorage
    localStorage.setItem("dark-mode", darkMode ? "false" : "true");

    // Tạo và phát sự kiện darkModeChange
    const darkModeEvent = new CustomEvent('darkModeChange', { detail: { darkMode: !darkMode } });
    document.dispatchEvent(darkModeEvent);

});



getUserProfile(access_token)
    .then((res) => {
        // const 
        if(res._id){
            const userName = res.name
            loginSelector.style.display = 'none'
            // logoutSelector.style.display=''
            logoutSelector.style.display = 'flex'
            helloUserSelector.innerHTML = `Xin chào ${userName}`
        }
    })
    .catch(() => {
        console.log("abc")
    })