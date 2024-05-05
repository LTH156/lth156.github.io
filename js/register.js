const registerHandler = async (fullname, email, password, confirm_password) => {
    try {
        const result = await fetch("https://novel-server-1.onrender.com/users/register", {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                name: fullname,
                email,
                password,
                confirm_password
            })
        });
        const data =  result.json();
        // console.log(data);
        return data; // Trả về dữ liệu từ máy chủ
    } catch (error) {
        console.error('Error:', error);
    }
};
function validatePassword(password) {
    const minLength = 8;
    const maxLength = 20; // Giả sử độ dài tối đa là 20 ký tự
    // Kiểm tra độ dài của mật khẩu
    if (password.length < minLength || password.length > maxLength) {
        return 'Mật khẩu phải chứa từ ' + minLength + ' đến ' + maxLength + ' ký tự.';
    }

    // Kiểm tra xem mật khẩu có chứa ít nhất một chữ hoa, một chữ thường, một số và một ký tự đặc biệt không
    const lowercaseRegex = /[a-z]/;
    const uppercaseRegex = /[A-Z]/;
    const digitRegex = /\d/;
    const specialCharRegex = /[@$!%*?&]/;

    if (!lowercaseRegex.test(password) || !uppercaseRegex.test(password) || !digitRegex.test(password) || !specialCharRegex.test(password)) {
        return 'Mật khẩu phải chứa ít nhất một chữ hoa, một chữ thường, một số và một ký tự đặc biệt.';
    }

    // Mật khẩu hợp lệ
    return null;
}
document.addEventListener('DOMContentLoaded', function () {
    const form = document.querySelector('form');
    form.addEventListener('submit', async function (event) {
        event.preventDefault(); // Ngăn chặn việc gửi form mặc định

        const fullname = form.elements['fullname'].value;
        const email = form.elements['email'].value;
        const password = form.elements['password'].value;
        const confirm_password = form.elements['confirm_password'].value;

        // Kiểm tra mật khẩu có đủ mạnh không


        // Sử dụng hàm validatePassword để kiểm tra mật khẩu
        const validationResult = validatePassword(password);

        if (validationResult) {
            alert(validationResult);
        } else {
            // Kiểm tra mật khẩu và xác nhận mật khẩu có giống nhau không
            if (password !== confirm_password) {
                alert('Mật khẩu và xác nhận mật khẩu không giống nhau!');
                return;
            }

            // Gọi hàm registerHandler để gửi yêu cầu đăng ký
            const result = await registerHandler(fullname, email, password, confirm_password);
            if (result && result.err && result.err.email) {
                alert('Email đã tồn tại!'); // Hiển thị thông báo khi email đã tồn tại
            }

            if (result.message === 'register success!') {
                alert("Chúc mừng đã tạo tài khoản thành công!");
                localStorage.setItem("access_token", result.result.access_token);
                localStorage.setItem("refresh_token", result.result.refresh_token);
                window.location.href = "./login.html";
            }
        }
    })
}); 