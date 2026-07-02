// ======================================
// CẤU HÌNH
// ======================================

const FILE_URL =
"http://biacoloa.local/wp-content/uploads/2026/07/WordPress_BaiViet_Bia.xlsx";

const POSTS_PER_PAGE = 5;

// ======================================
// BIẾN TOÀN CỤC
// ======================================

let posts = [];
let filteredPosts = [];
let currentPage = 1;

// ======================================
// KHỞI TẠO
// ======================================

document.addEventListener("DOMContentLoaded", init);

async function init() {

    await loadExcel();

    filteredPosts = [...posts];

    renderPosts();

    renderPagination();

    initSearch();

}

// ======================================
// ĐỌC FILE EXCEL
// ======================================

async function loadExcel() {

    try {

        const response = await fetch(FILE_URL);

        if (!response.ok) {
            throw new Error("Không đọc được Excel");
        }

        const buffer = await response.arrayBuffer();

        const workbook = XLSX.read(buffer, {
            type: "array"
        });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        posts = XLSX.utils.sheet_to_json(sheet);

        posts = posts.filter(post => post["Mã bài viết"]);

    }

    catch(error){

        console.log(error);

    }

}

// ======================================
// HIỂN THỊ DANH SÁCH
// ======================================

function renderPosts(){

    const newsList=document.getElementById("news-list");

    if(!newsList) return;

    const start=(currentPage-1)*POSTS_PER_PAGE;

    const end=start+POSTS_PER_PAGE;

    const pagePosts=filteredPosts.slice(start,end);

    let html="";

    pagePosts.forEach(post=>{

        html+=`

        <div class="news-card">

            <a href="${post["source_url"]}">

                <img src="${post["featured_image"]}" alt="">

            </a>

            <div class="content">

                <h3>

                    <a href="${post["source_url"]}">

                        ${post["post_title (trang chi tiet)"]}

                    </a>

                </h3>

                <small>

                    <a
                        href="${post["link_trang"]}"
                        target="_blank"
                        rel="noopener noreferrer"
                        class="copyright-link"
                    >
                        © HEINEKEN Việt Nam
                    </a>

                    &nbsp; | &nbsp;

                    ${post["post_category"]}

                </small>

                <p>

                    ${post["excerpt (support cho title)"]}

                </p>

                <a

                    class="read-more"

                    href="${post["source_url"]}"

                >

                    XEM CHI TIẾT

                </a>

            </div>

        </div>

        `;

    });

    if(html===""){

        html="<p>Không có bài viết.</p>";

    }

    newsList.innerHTML=html;

}

// ======================================
// PHÂN TRANG
// ======================================

function renderPagination(){

    const pagination=document.getElementById("pagination");

    if(!pagination) return;

    const totalPages=Math.ceil(filteredPosts.length/POSTS_PER_PAGE);

    let html="";

    if(currentPage>1){

        html+=`
        <a href="#" onclick="changePage(${currentPage-1});return false;">
            «
        </a>
        `;

    }

    for(let i=1;i<=totalPages;i++){

        html+=`

        <a

            href="#"

            class="${i===currentPage?"active":""}"

            onclick="changePage(${i});return false;">

            ${i}

        </a>

        `;

    }

    if(currentPage<totalPages){

        html+=`
        <a href="#" onclick="changePage(${currentPage+1});return false;">
            »
        </a>
        `;

    }

    pagination.innerHTML=html;

}

// ======================================
// CHUYỂN TRANG
// ======================================

function changePage(page){

    currentPage=page;

    renderPosts();

    renderPagination();

    window.scrollTo({

        top:0,

        behavior:"smooth"

    });

}

// ======================================
// TÌM KIẾM
// ======================================

function initSearch(){

    const input=document.getElementById("searchInput");

    if(!input) return;

    input.addEventListener("input",function(){

        const keyword=this.value.toLowerCase().trim();

        if(keyword===""){

            filteredPosts=[...posts];

        }else{

            filteredPosts=posts.filter(post=>{

                return String(post["post_title (trang chi tiet)"]||"")

                .toLowerCase()

                .includes(keyword);

            });

        }

        currentPage=1;

        renderPosts();

        renderPagination();

    });

}

// ======================================
// CHUYỂN NGÀY EXCEL
// ======================================

function excelDateToString(excelDate){

    if(!excelDate) return "";

    if(typeof excelDate==="string"){

        return excelDate;

    }

    const date=XLSX.SSF.parse_date_code(excelDate);

    return `${String(date.d).padStart(2,"0")}/${String(date.m).padStart(2,"0")}/${date.y}`;

}
