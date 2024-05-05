const loginSelector = document.querySelector('.login');
const registerSelector = document.querySelector('.register')


const loginHandler = async (email, password) => {
    try {
        const result = await fetch("https://novel-server-1.onrender.com/users/login", {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                email,
                password
            })
        });
        const data = await result.json();
        return data; // Trả về dữ liệu từ máy chủ
    } catch (error) {
        console.error('Error:', error);
    }
}

document.addEventListener('DOMContentLoaded', function () {
    const form = document.querySelector('form');
    form.addEventListener('submit', async function (event) {
        event.preventDefault(); // Ngăn chặn việc gửi form mặc định
        const username = form.elements['username'].value;
        const password = form.elements['password'].value;

        // Gọi hàm loginHandler để gửi yêu cầu đăng nhập
        const result = await loginHandler(username, password);

        // Xử lý kết quả đăng nhập
        if (result && result.message === "login success!") {
            alert('Đăng nhập thành công!');
            localStorage.setItem("access_token", result.result)
            // Lưu token vào localStorage hoặc cookie
            // Chuyển hướng đến trang sau khi đăng nhập thành công
            window.location.href = "./index.html";
        } else {
            alert('Đăng nhập thất bại! Vui lòng kiểm tra lại thông tin đăng nhập.');
        }
    });
});
