var DoneAll = [];

var normal = document.getElementById("normal");
var advanced = document.getElementById("advanced");
var normalCon = document.getElementById("normalCon");
var advancedCon = document.getElementById("advancedCon");
var tobo = document.querySelector(".table_group");
normal.addEventListener("click", function () {
    if (!(normal.classList.contains("disabled"))) {
        advancedCon.style.display = "none";
        normalCon.style.display = "block";
        normal.classList.add("disabled");
        advanced.classList.remove("disabled");
    }
})
advanced.addEventListener("click", function () {
    if (!(advanced.classList.contains("disabled"))) {
        advancedCon.style.display = "block";
        normalCon.style.display = "none";
        normal.classList.remove("disabled");
        advanced.classList.add("disabled");

    }
})

function generateLegacyId() {
    const timestamp = Date.now().toString(36); 
    // console.log(Date.now().toString(36))
    const randomStr = Math.random().toString(36).substring(2, 8);
    return `${timestamp}-${randomStr}`;
}

/////////////////////////////////////////
function normalizeArabicText(text) {
    if (!text) return "";
    // Convert Arabic-Indic numerals (٠١٢٣٤٥٦٧٨٩) to Western numerals (0123456789)
    let normalized = text.replace(/[٠-٩]/g, (d) => '٠١٢٣٤٥٦٧٨٩'.indexOf(d));
    // Normalize Unicode (NFC form)
    normalized = normalized.normalize("NFC");
    // Trim and collapse multiple spaces into one
    normalized = normalized.trim().replace(/\s+/g, ' ');
    return normalized;
}

document.getElementById('excelFile').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const fullOk = await isXlsxFile(file);
    if (!fullOk) {
        alert('الملف ليس بصيغة XLSX صحيحة.');
        return;
    }


    const reader = new FileReader();
    reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, {
            type: 'array', cellDates: true, cellText: false, cellNF: true
        });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: false, dateNF: 'yyyy-mm-dd' });
        // console.log(json)
        let testDate = [];
        json.forEach(row => {
            if (!isNaN(row["م"])) {
                let rowSubtopic = normalizeArabicText(row["الموضوع الفرعي"]);
                let topicInfoId = DataAD.find((t) => normalizeArabicText(t.Subtopic) === rowSubtopic);
                if (topicInfoId) {

                    let [day, month, year] = row["التاريخ"].split("/");
                    year = Number(year) < 2000 ? Number(year) + 2000 : Number(year);

                    testDate.push({
                        year: year,
                        month: Number(month),
                        day: Number(day),

                        men: Number(row["ذكور"]),
                        child: Number(row["اطفال"]),
                        women: Number(row["إناث"]),

                        MainTopic: row["الموضوع الرئيسي"],
                        Subtopic: row["الموضوع الفرعي"],

                        id: topicInfoId.id,

                        uid:generateLegacyId() ,

                        in: row["خارجية"] === "داخلية",
                        out: row["خارجية"] === "خارجية",
                        counter: Number(row["م"]),
                        dayWeek: row["اليوم"],
                        seminarCount: row["عدد الندوات"] ? Number(row["عدد الندوات"]) : 1,
                    });
                }
            }
        })

        if (testDate.length > 0) {
            DoneAll = [...testDate];
            drowTopic(sortTopic(DoneAll));
        }
        // console.log(testDate)

    };
    reader.readAsArrayBuffer(file);
});

function isXlsxFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = (e) => {
            const arr = new Uint8Array(e.target.result).subarray(0, 4);
            const header = Array.from(arr).map(b => b.toString(16).padStart(2, '0')).join('');
            resolve(header === '504b0304');
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file.slice(0, 4));
    });
}

/////////////////////////////////////////


var tablegroupdivider = document.querySelector(".table-group-divider");
let ans = false;
window.onload = function () {
    if (localStorage.getItem("saveDataForLate")) {
        // console.log(localStorage.getItem("saveDataForLate"))
        ans = confirm("هناك ملفات مخزنة من قبل هل تريد استرجاعها لتكملة التعديل عليها ؟\n لاحظ : في حالة الرفض سوف يتم حذف هذة البيانات القديمة من المخزن.");

        if (ans) {
            DoneAll = JSON.parse(localStorage.getItem("saveDataForLate"));
            // console.log(DoneAll);
            // generateFunc();
            drowTopic(sortTopic(DoneAll));
        } else
            localStorage.removeItem("saveDataForLate");
    }

    // console.log(ans);

    /////////////////////////
    DataAD.forEach((element) => {
        tablegroupdivider.innerHTML += `
        <tr>
            <td>${element.id}</td>
            <td>${element.MainTopic}</td>
            <td>${element.Subtopic}</td>
            <td><input type="checkbox" name="" id="${element.id}in" class="checkBox" onclick="inCheck(\`${element.id}in\`)"></td>
            <td><input type="checkbox" name="" id="${element.id}out"  class="checkBox" onclick="outCheck(\`${element.id}out\`)"></td>
            <td><input type="number" name="" id="${element.id}child"  class="numbIn" onblur="childNe(\`${element.id}child\`)"></td>
            <td><input type="number" name="" id="${element.id}man"  class="numbIn" onblur="manNe(\`${element.id}man\`)"></td>
            <td><input type="number" name="" id="${element.id}woman"  class="numbIn" onblur="womanNe(\`${element.id}woman\`)"></td>
            <td><input type="number" name="" id="${element.id}seminar" class="numbIn" min="1" value="1" onblur="seminarNe(\`${element.id}seminar\`)"></td>
        </tr>
        `
    })
}
//////////////////////////////////////////////////
function childNe(e) {
    let val = document.getElementById(e).value;
    if (val) {
        DataAD[parseInt(e) - 1].child = +(val);
        // console.log(parseInt(e) - 1);
        // console.log(val);

    }
}
function manNe(e) {
    let val = document.getElementById(e).value;
    if (val) {
        DataAD[parseInt(e) - 1].men = +(val);
        // console.log(parseInt(e) - 1);
        // console.log(val);

    }
}
function womanNe(e) {
    let val = document.getElementById(e).value;
    if (val) {
        DataAD[parseInt(e) - 1].women = +(val);
        // console.log(parseInt(e) - 1);
        // console.log(val);

    }
}
function seminarNe(e) {
    let val = document.getElementById(e).value;
    if (val) {
        DataAD[parseInt(e) - 1].seminarCount = +(val);
    }
}
//////////////////////////////////////////////////
var child = 0, man = 0, woman = 0;
var womenInput = document.getElementById("women");
var menInput = document.getElementById("men")
var childInput = document.getElementById("child");

document.querySelector(".done").onclick = function () {
    if (!(document.querySelector(".done").classList.contains("disabled"))) {
        child = +(childInput.value);
        man = +(menInput.value);
        woman = +(womenInput.value);
        DataAD.forEach(ele => {
            ele.child = child;
            ele.men = man;
            ele.women = woman;
        })
        // console.log(DataAD);
        childInput.setAttribute("disabled", "");
        menInput.setAttribute("disabled", "");
        womenInput.setAttribute("disabled", "");
        document.querySelector(".done").classList.add("disabled");
        document.querySelector(".reset").classList.remove("disabled");
    }
}

document.querySelector(".reset").onclick = function () {
    if (!(document.querySelector(".reset").classList.contains("disabled"))) {
        // console.log(child + " " + man + " " + woman);
        document.getElementById("child").removeAttribute("disabled");
        document.getElementById("men").removeAttribute("disabled");
        document.getElementById("women").removeAttribute("disabled");
        // document.getElementById("child").value = "";
        // document.getElementById("men").value = "";
        // document.getElementById("women").value = "";
        document.querySelector(".done").classList.remove("disabled");
        document.querySelector(".reset").classList.add("disabled");
    }
}
////////////////////////////////////////////////

function inCheck(n) {
    let test = document.getElementById(n);
    let num = parseInt(n);
    let botion = DataAD.findIndex(ele => ele.id === num);
    // console.log(test.checked)
    if (test.checked) {
        test.parentElement.parentElement.classList.add("in");
        DataAD[botion].in = true;

    } else {
        test.parentElement.parentElement.classList.remove("in");
        DataAD[botion].in = false;
    }
    // console.log(DataAD);
    bothCheck(n);
}
function outCheck(n) {
    let test = document.getElementById(n);
    let num = parseInt(n);
    let botion = DataAD.findIndex(ele => ele.id === num);
    if (test.checked) {
        test.parentElement.parentElement.classList.add("out");
        DataAD[botion].out = true;
    } else {
        test.parentElement.parentElement.classList.remove("out");
        DataAD[botion].out = false;
    }
    bothCheck(n);
    // console.log(DataAD);
}
function bothCheck(n) {
    let test = document.getElementById(n);
    if (test.parentElement.parentElement.classList.contains("in") && test.parentElement.parentElement.classList.contains("out")) {
        test.parentElement.parentElement.classList.add("both");
    } else {
        test.parentElement.parentElement.classList.remove("both");
    }
}
////////////////////////////////////////////////

var daysCon = {
    in: [],
    out: [],
    bothINday: "",
};
var campaign365Days = {}; // {dayNumber: "in" or "out"}
var startDateInput = document.getElementById("startDate");
var endDateInput = document.getElementById("endDate");
var startDate, endDate;
startDate = new Date(startDateInput.value);
endDate = new Date(endDateInput.value);
var dataSelectIN = document.querySelectorAll("#dataSelect .inWeek .dayInfo");
var dataSelectOUT = document.querySelectorAll("#dataSelect .outWeek .dayInfo");
var dataSelectCampaign365 = document.querySelectorAll("#dataSelect .campaign365Week .dayInfo");
var generate = document.querySelector(".generate");
var bothINday = document.getElementById("bothINday");
var supDate = document.querySelector(".supDate");
var inandoutChose = document.querySelectorAll(".bothInDay")
// console.log(inandoutChose)
supDate.onclick = () => {
    daysCon = {
        in: {},
        out: {},
        bothINday: "",
    };
    campaign365Days = {};
    startDate = new Date(startDateInput.value);
    endDate = new Date(endDateInput.value);

    dataSelectIN.forEach((e) => {
        if (e.querySelector("input[type=checkbox]").checked) {
            daysCon.in[+(e.querySelector("input[type=checkbox]").value)] = +e.querySelector("input[type=number]").value;
        }
    });

    dataSelectOUT.forEach((e) => {
        if (e.querySelector("input[type=checkbox]").checked) {
            daysCon.out[+(e.querySelector("input[type=checkbox]").value)] = +e.querySelector("input[type=number]").value;
        }
    })

    // حفظ إعدادات حملة 365 يوم سلامة لكل يوم
    dataSelectCampaign365.forEach((e) => {
        if (e.querySelector("input[type=checkbox]").checked) {
            let dayNum = +(e.querySelector("input[type=checkbox]").value);
            let type = e.querySelector("select").value; // "in" or "out"
            campaign365Days[dayNum] = type;
        }
    });

    generate.classList.remove("disabled");
}
////////////////////////////////////////////////////////////////////
var monthNames = [
    "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
    "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"
];
var daysOfWeek = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"];

function sortTopic(array) {
    // console.log(array);
    localStorage.setItem("saveDataForLate", JSON.stringify(array));
    const x = {};
    array.sort((a, b) => {
        if (a.year !== b.year) {
            return a.year - b.year;
        }
        if (a.month !== b.month) {
            return a.month - b.month;
        }
        return a.day - b.day;
    });

    array.forEach((element) => {
        if (x[`${element.month}-${element.year}`]) {
            x[`${element.month}-${element.year}`].push(element);
        } else
            x[`${element.month}-${element.year}`] = [element];
    });
    let n =1;
    // console.log(array)
    array.forEach((e,i)=>{
        // console.log(e)
        if( array[i].month>1 && i>0 && array[i].month!=array[i-1].month){n=1;}
        e.counter=n++
    })
    return x;
}

function drowTopic(object) {
    // console.log("drowTopic")
    // tobo.innerHTML = "";
    tobo.innerHTML = ` 
    <thead class="table-group-divider">
    <th>م</th>
    <th>التاريخ</th>
    <th>اليوم</th>
    <th>الموضوع الرئيسي</th>
    <th>الموضوع الفرعي</th>
    <th>خارجية</th>
    <th>داخلية</th>
    <th>اطفال</th>
    <th>ذكور</th>
    <th>إناث</th>
    <th>عدد الندوات</th>
    <th style ="width : 7%;">تعديل</th>
    <th style ="width : 7%;">حذف</th>
    </thead>
    <tbody id="tbody">
    <!-- ////// -->
    </tbody>`;
    var tbody = document.getElementById("tbody");
    Object.keys(object).forEach(ele => {
        let detti = ele.split("-");
        tbody.innerHTML += `<tr>
        <td colspan="13" class="titelTopics fw-bolder">الجلسات التثقيفية عن شهر ${detti[0]} لعام ${detti[1]}</td>
        </tr>
        `
        object[ele].forEach(i => {
            // console.log(i)
            let formattedDate = i.day + "/" + i.month + "/" + i.year;
            tbody.innerHTML += `          
            <tr class = ${i.in ? "inter" : "outter"}>
            <td>${i.counter}</td>
            <td>${formattedDate}</td>
            <td>${i.dayWeek}</td>
            <td>${i.MainTopic}</td>
            <td>${i.Subtopic}</td>
            <td colspan="2" class="fw-bold">${i.in ? "داخلية" : "خارجية"}</td>
            <td>${i.child}</td>
            <td>${i.men}</td>
            <td>${i.women}</td>
            <td>${i.seminarCount || 1}</td>
            <td class ="${i.cahnged ? "retopict" : ""}"><i class="fa-solid fa-recycle reTopic" onclick = "reTopic(\'${i.uid}\')"></i></td>
            <td><i class="fa-solid fa-trash-can deleteTopic" onclick = "deleteTopic(\'${i.uid}\')"></i></td>
            </tr>
            `
        })
        // console.log(ele)
    })
    tbody.innerHTML += `          
    <tr class = "NewTopicL" onclick = "AddNewTopic()">
    <td colspan = "13"><i class="fa-solid fa-circle-plus"></i></td>
    </tr>
    `;
    document.querySelector(".topicContan").style.display = "block";
    generateMonthYear.classList.remove("disabled");
}





var generateMonthYear = document.querySelector(".generateMonthYear");


generate.addEventListener("click", generateFunc)

function generateFunc() {
    var counter = 1;
    DoneAll = [];

    // tbody.innerHTML = "";
    // var tbody = document.getElementById("tbody");
    let helper = new Date(startDate);
    let especifcIN = DataAD.filter(ele => ele.in);
    let especifcOut = DataAD.filter(ele => ele.out);
    var directDetiIn = (normal.classList.contains("disabled")) ? DataAD : especifcIN;
    var directDetiOut = (normal.classList.contains("disabled")) ? DataAD : especifcOut;
    let maxRandomIn = directDetiIn.length;
    let maxRandomOut = directDetiOut.length;
    // console.log(directDetiIn);
    // let currentMon = startDate.getMonth() + 1;
    // let currentDay = startDate.getDate();
    // let currentYear = startDate.getFullYear();
    while (startDate <= endDate) {


        var day = startDate.getDate();
        var month = startDate.getMonth() + 1; // يضاف واحد لأن الشهور تبدأ من 0
        var year = startDate.getFullYear();
        var formattedDate = day + "/" + month + "/" + year;
        let checkIn = daysCon.in[startDate.getDay()];
        let checkOut = daysCon.out[startDate.getDay()];
        if (checkIn) {
            for (let index = 0; index < checkIn; index++) {
                let randomNumber = Math.floor(Math.random() * (maxRandomIn - 0 + 0)) + 0;
                // let trueObg = directDetiIn[randomNumber];
                ////////////////////////////////////////////
                var newop = {
                    year: year,
                    month: month,
                    day: day,
                    men: directDetiIn[randomNumber].men,
                    child: directDetiIn[randomNumber].child,
                    women: directDetiIn[randomNumber].women,
                    MainTopic: directDetiIn[randomNumber].MainTopic,
                    Subtopic: directDetiIn[randomNumber].Subtopic,
                    id: directDetiIn[randomNumber].id,
                    in: true,
                    out: false,
                    counter: counter,
                    dayWeek: daysOfWeek[startDate.getDay()],
                    seminarCount: directDetiIn[randomNumber].seminarCount || 1,
                    uid: generateLegacyId() 
                }
                DoneAll.push(newop);
                counter++;
            }
        }

        if (checkOut) {
            for (let index = 0; index < checkOut; index++) {
                let randomNumber = Math.floor(Math.random() * (maxRandomOut - 0 + 0)) + 0;
                ////////////////////////////////////////////
                var newop = {
                    year: year,
                    month: month,
                    day: day,
                    men: directDetiOut[randomNumber].men,
                    child: directDetiOut[randomNumber].child,
                    women: directDetiOut[randomNumber].women,
                    MainTopic: directDetiOut[randomNumber].MainTopic,
                    Subtopic: directDetiOut[randomNumber].Subtopic,
                    id: directDetiOut[randomNumber].id,
                    in: false,
                    out: true,
                    counter: counter,
                    dayWeek: daysOfWeek[startDate.getDay()],
                    seminarCount: directDetiOut[randomNumber].seminarCount || 1,
                    uid: generateLegacyId() 
                }
                DoneAll.push(newop);
                counter++;
            }
        }

        // إضافة "حملة 365 يوم سلامة" حسب إعدادات كل يوم
        let campaignType = campaign365Days[startDate.getDay()];
        if (campaignType) {
            let campaign365Data = DataAD.find(t => t.id === 53);
            // أخذ قيم الأطفال والذكور والإناث من آخر ندوة تم إنشاؤها في نفس اليوم
            let lastEntry = DoneAll[DoneAll.length - 1];
            if (campaign365Data) {
                var campaignEntry = {
                    year: year,
                    month: month,
                    day: day,
                    men: lastEntry ? lastEntry.men : 0,
                    child: lastEntry ? lastEntry.child : 0,
                    women: lastEntry ? lastEntry.women : 0,
                    MainTopic: campaign365Data.MainTopic,
                    Subtopic: campaign365Data.Subtopic,
                    id: campaign365Data.id,
                    in: campaignType === "in",
                    out: campaignType === "out",
                    counter: counter,
                    dayWeek: daysOfWeek[startDate.getDay()],
                    seminarCount: lastEntry ? (lastEntry.seminarCount || 1) : 1,
                    uid :generateLegacyId() 
                }
                DoneAll.push(campaignEntry);
                counter++;
            }
        }

        startDate.setDate(startDate.getDate() + 1); // يزيد التاريخ بيوم واحد
    }

    // console.log(sortTopic(DoneAll));
    drowTopic(sortTopic(DoneAll))

    startDate = helper;

    ////////////////////////////////////////////////
    // console.log(DoneAll);

}

function deleteTopic(e) {
    // console.log(dataMontlyAll)
    // console.log(e)
    let ind = DoneAll.findIndex(ele => ele.uid == e);
    DoneAll.splice(ind, 1);
    // console.log(DoneAll)
    drowTopic(sortTopic(DoneAll));
}
/////////////////////////////////////////////////////////////////////
var overLay = document.querySelector(".overLay");
var Save_overLay = document.querySelector(".Save_overLay");
var Close_overLay = document.querySelector(".Close_overLay");
function reTopic(e) {
    let ind = DoneAll.findIndex(ele => ele.uid == e);
    let ele = DoneAll[ind];
    // console.log(DoneAll);


    let oldDeta = `${ele.year}-${ele.month < 10 ? `0${ele.month}` : ele.month}-${ele.day < 10 ? `0${ele.day}` : ele.day}`;
    let startDetaLimit = `${startDate.getFullYear()}-${(startDate.getMonth() + 1) < 10 ? `0${(startDate.getMonth() + 1)}` : (startDate.getMonth() + 1)}-${(startDate.getDate()) < 10 ? `0${(startDate.getDate())}` : (startDate.getDate())}`;
    let endDetaLimit = `${endDate.getFullYear()}-${(endDate.getMonth() + 1) < 10 ? `0${(endDate.getMonth() + 1)}` : (endDate.getMonth() + 1)}-${(endDate.getDate()) < 10 ? `0${(endDate.getDate())}` : (endDate.getDate())}`;
    // console.log(oldDeta);

    const uniqueSubtopics = [...new Set(DataAD.map(item => item.MainTopic))];
    // console.log(uniqueSubtopics);

    overLay.innerHTML = "";
    overLay.innerHTML += `<div class="befor-element" onclick="closeLay()"></div>`;
    overLay.innerHTML += `
    <div class="inputs-contain">
    <p class="fw-bold m-0">المسلسل :${ele.counter}</p>
    <div class="form1">
      <div><p for="startDate" class="m-0">التاريخ</p><input type="date" id="startDateTime" name="startDate" value="${oldDeta}" max="${endDetaLimit}" min="${startDetaLimit}"></div>
      <div><p for="startDate" class="m-0">الموضوع الرئيسي</p><select name="" id="MainTopicEsp" onchange = "supTop()"><option value=""></option></select></div>
      <div><p for="startDate" class="m-0">الموضوع الفرعي</p><select name="" id="SupTopicEsp" ><option value=""></option></select></div>
      <div><p for="startDate" class="m-0">مكان الندوة</p><select name="" id="INoutCONDETION" value = "">
      <option value="خارجية">خارجية</option>
      <option value="داخلية">داخلية</option>
      </select></div>
      <div><p for="startDate" class="m-0">أطفال</p><input type="number" name="" id="childchange" min="0" max="1000" value ="${ele.child}"></div>
      <div><p for="startDate" class="m-0">ذكور</p><input type="number" name="" id="manchange" min="0" max="1000" value ="${ele.men}"></div>
      <div><p for="startDate" class="m-0">إناث</p><input type="number" name="" id="womanchange" min="0" max="1000" value ="${ele.women}"></div>
      <div><p for="startDate" class="m-0">عدد الندوات</p><input type="number" name="" id="seminarCountChange" min="1" max="1000" value ="${ele.seminarCount || 1}"></div>

    </div>
    <div class="bottonssd d-flex justify-content-start ">
      <button type="button" class="btn btn-primary mx-2 Save_overLay"onclick = "saveChangesLay(${ind})">حفظ</button>
      <button type="button" class="btn btn-danger mx-2 Close_overLay" onclick = "closeLay()">اغلاق</button>
    </div>
    </div>
  `
    var MainTopicEsp = document.getElementById("MainTopicEsp");
    MainTopicEsp.innerHTML = ``;
    uniqueSubtopics.forEach(e => {
        MainTopicEsp.innerHTML += `<option value="${e}">${e}</option>`;
    })
    MainTopicEsp.value = ele.MainTopic;
    let SupTopicEsp = document.getElementById("SupTopicEsp")
    supTop();
    SupTopicEsp.value = ele.Subtopic;
    let INoutCONDETION = document.getElementById("INoutCONDETION")
    INoutCONDETION.value = ele.in ? "داخلية" : "خارجية";

    overLay.style.display = "flex";
}
function supTop() {
    let SupTopicEsp = document.getElementById("SupTopicEsp")
    let NowOn = MainTopicEsp.value;
    let Allsup = DataAD.filter(e => e.MainTopic == NowOn)
    SupTopicEsp.innerHTML = '';
    Allsup.forEach(e => {
        SupTopicEsp.innerHTML += `<option value="${e.Subtopic}">${e.Subtopic}</option>`;
    })
    SupTopicEsp.value = "";
}
// console.log(Close_overLay)
function saveChangesLay(params) {
    // console.log(params)
    // console.log(DoneAll)
    let realIndex = DoneAll.findIndex(ele => ele.counter == params);
    let SupTopicSave = document.getElementById("SupTopicEsp");
    if (SupTopicSave.value) {
        let dateTime = new Date(document.getElementById("startDateTime").value);
        var dayo = dateTime.getDate();
        DoneAll[params].day = dayo;
        var montho = dateTime.getMonth() + 1; // يضاف واحد لأن الشهور تبدأ من 0
        DoneAll[params].month = montho;
        var yearo = dateTime.getFullYear();
        DoneAll[params].year = yearo;
        var formattedDateo = dayo + "/" + montho + "/" + yearo;
        var dayweeko = daysOfWeek[dateTime.getDay()];
        DoneAll[params].dayWeek = dayweeko;
        DoneAll[params].cahnged = true;
        // console.log(formattedDateo);
        // console.log(dayweeko);

        let MainTopicSave = document.getElementById("MainTopicEsp").value;
        DoneAll[params].MainTopic = MainTopicSave;

        DoneAll[params].Subtopic = SupTopicSave.value;
        let INoutCONDETION = document.getElementById("INoutCONDETION").value;
        if (INoutCONDETION == "داخلية") {
            DoneAll[params].in = true;
            DoneAll[params].out = false;
        } else if (INoutCONDETION == "خارجية") {
            DoneAll[params].in = false;
            DoneAll[params].out = true;
        }

        let childchange = Number(document.getElementById("childchange").value);
        let manchange = Number(document.getElementById("manchange").value);
        let womanchange = Number(document.getElementById("womanchange").value);
        let seminarCountChange = Number(document.getElementById("seminarCountChange").value) || 1;
        DoneAll[params].child = childchange;
        DoneAll[params].men = manchange;
        DoneAll[params].women = womanchange;
        DoneAll[params].seminarCount = seminarCountChange;
        DoneAll[params].id = DataAD.find(e => e.Subtopic == SupTopicSave.value).id;

        drowTopic(sortTopic(DoneAll));

        overLay.style.display = "none";
    } else {
        alert("من فضلك تأكد من اختيار الموضوع الفرعي");
        SupTopicSave.style.border = "2px solid red ";
    }
}
//////////////////////////////////////////////////////////////////////////////////
function AddNewTopic() {
    // console.log(6);
    let startDetaLimit = `${startDate.getFullYear()}-${(startDate.getMonth() + 1) < 10 ? `0${(startDate.getMonth() + 1)}` : (startDate.getMonth() + 1)}-${(startDate.getDate()) < 10 ? `0${(startDate.getDate())}` : (startDate.getDate())}`;
    let endDetaLimit = `${endDate.getFullYear()}-${(endDate.getMonth() + 1) < 10 ? `0${(endDate.getMonth() + 1)}` : (endDate.getMonth() + 1)}-${(endDate.getDate()) < 10 ? `0${(endDate.getDate())}` : (endDate.getDate())}`;
    // let max = 0;
    // // console.log(DoneAll);
    // for (let i of DoneAll) {
    //     if (i.counter > max) {
    //         max = i.counter;
    //     }
    // }
    // max++;
    overLay.innerHTML = "";
    overLay.innerHTML += `<div class="befor-element" onclick="closeLay()"></div>`;
    overLay.innerHTML += `
    <div class="inputs-contain">
    <p class="fw-bold m-0">المسلسل :${0}</p>
    <div class="form1">
      <div><p for="startDate" class="m-0">التاريخ</p><input type="date" id="startDateTime" name="startDate"  max="${endDetaLimit}" min="${startDetaLimit}"></div>
      <div><p for="startDate" class="m-0">الموضوع الرئيسي</p><select name="" id="MainTopicEsp" onchange = "supTop()"><option></option></select></div>
      <div><p for="startDate" class="m-0">الموضوع الفرعي</p><select name="" id="SupTopicEsp" ><option value=""></option></select></div>
      <div><p for="startDate" class="m-0">مكان الندوة</p><select name="" id="INoutCONDETION" value = "">
      <option value="" hidden></option>
      <option value="خارجية">خارجية</option>
      <option value="داخلية">داخلية</option>
      </select></div>
      <div><p for="startDate" class="m-0">أطفال</p><input type="number" name="" id="childchange" min="0" max="1000" ></div>
      <div><p for="startDate" class="m-0">ذكور</p><input type="number" name="" id="manchange" min="0" max="1000" ></div>
      <div><p for="startDate" class="m-0">إناث</p><input type="number" name="" id="womanchange" min="0" max="1000" ></div>
      <div><p for="startDate" class="m-0">عدد الندوات</p><input type="number" name="" id="seminarCountChange" min="1" max="1000" value="1"></div>

    </div>
    <div class="bottonssd d-flex justify-content-start ">
      <button type="button" class="btn btn-primary mx-2 Save_overLay"onclick = "acceptNewTopic()">حفظ</button>
      <button type="button" class="btn btn-danger mx-2 Close_overLay" onclick = "closeLay()">اغلاق</button>
    </div>
    </div>
  `;
    const uniqueSubtopics = [...new Set(DataAD.map(item => item.MainTopic))];
    var MainTopicEsp = document.getElementById("MainTopicEsp");
    MainTopicEsp.innerHTML = `<option value="" hidden></option>`;
    // console.log(uniqueSubtopics)
    uniqueSubtopics.forEach(e => {
        MainTopicEsp.innerHTML += `<option value="${e}">${e}</option>`;
    })
    // console.log(uniqueSubtopics)

    overLay.style.display = "block";
}

function acceptNewTopic() {

    let dateN = new Date(document.getElementById("startDateTime").value);
    let day = dateN.getDate();
    let month = dateN.getMonth() + 1; // يضاف واحد لأن الشهور تبدأ من 0
    let year = dateN.getFullYear();
    let formattedDateN = day + "/" + month + "/" + year; // التاريخ

    let dayWeaka = daysOfWeek[dateN.getDay()]; // اليوم في الأسبوع
    // console.log(formattedDateN, dayWeaka);

    let INoutCONDETION = document.getElementById("INoutCONDETION").value;


    let MainTopicEsp = document.getElementById("MainTopicEsp").value;
    let SupTopicSave = document.getElementById("SupTopicEsp").value;

    let childchange = Number(document.getElementById("childchange").value);
    let manchange = Number(document.getElementById("manchange").value);
    let womanchange = Number(document.getElementById("womanchange").value);
    // console.log(childchange)


    if (dayWeaka && MainTopicEsp && SupTopicSave && INoutCONDETION) {
        let eleM = DataAD.find(e => e.MainTopic == MainTopicEsp && e.Subtopic == SupTopicSave);
        let seminarCountNew = Number(document.getElementById("seminarCountChange").value) || 1;
        var newop = {
            year: year,
            month: month,
            day: day,
            men: manchange,
            child: childchange,
            women: womanchange,
            MainTopic: MainTopicEsp,
            Subtopic: SupTopicSave,
            id: eleM.id,
            // in: false,
            // out: true,
            uid: generateLegacyId() ,
            counter: 0,
            dayWeek: dayWeaka,
            seminarCount: seminarCountNew,
        }
        if (INoutCONDETION == "داخلية") {
            newop.in = true;
            newop.out = false;
        } else if (INoutCONDETION == "خارجية") {
            newop.in = false;
            newop.out = true;
        }
        // console.log(newop);
        DoneAll.push(newop);
        ///////////////////////////////////////////////////////////////////////

        drowTopic(sortTopic(DoneAll));

        overLay.style.display = "none";
    } else {
        alert("من فضلك اكمل باقي البيانات");
        if (!dayWeaka) {
            document.getElementById("startDateTime").style.border = "2px solid red";
        }
        if (!MainTopicEsp) {
            document.getElementById("MainTopicEsp").style.border = "2px solid red";
        }
        if (!SupTopicSave) {
            document.getElementById("SupTopicEsp").style.border = "2px solid red";
        }
        if (!INoutCONDETION) {
            document.getElementById("INoutCONDETION").style.border = "2px solid red";
        }
    }

}
//////////////////////////////////////////////////////////////////////////////////
let closeLay = () => {
    overLay.style.display = "none";
}
/////////////////////////////////////////////////////////////////////

generateMonthYear.addEventListener("click", () => {

    var dataMontlyAll = sortTopic(DoneAll);
    // console.log(dataMontlyAll)

    //////////////////////////////////////////////////////////////////
    var monthly = document.querySelector(".monthly");
    var monthlyCon = document.querySelector(".monthlyCon");
    monthlyCon.classList.remove("d-none")
    monthly.innerHTML = "";
    Object.keys(dataMontlyAll).forEach(ele => {
        let item = dataMontlyAll[ele];
        let info = ele.split("-")
        // console.log(item, info);
        monthly.innerHTML += `
        <p class="mt-2 text-center fw-bold">خطة شهر ${info[0]} (${monthNames[+(info[0]) - 1]}) لعام ${info[1]}</p>
        `;
        monthly.innerHTML += `
        <table class="table-group-divider" value = "خطة شهر ${info[0]} (${monthNames[+(info[0]) - 1]}) لعام ${info[1]}">
        <thead>
          <tr>
            <th style="width: 15%;">الموضوع الرئيسي</th>
            <th style="width: 15%;">الموضوع الفرعي</th>
            <th style="width: 6%;">ندوة داخلية</th>
            <th style="width: 6%;">مشورة عامة</th>
            <th style="width: 6%;">ندوة خارجية</th>
            <th style="width: 6%;">لفاء جماهيري</th>
            <th style="width: 6%;">ندوات قوافل</th>
            <th style="width: 6%;">أطفال</th>
            <th style="width: 6%;">ذكور</th>
            <th style="width: 6%;">إناث</th>
            <th style="width: 6%;">استخدام وسائل إعلامية</th>
            <th style="width: 6%;">انشطة إعلامية رقمية</th>
            <th style="width: 10%;">ملاحظــات</th>
          </tr>
        </thead>
        <tbody>
        </tbody>
      </table>
      `;
        var BodyMonthly = document.querySelectorAll(".monthly tbody");
        BodyMonthly = BodyMonthly[BodyMonthly.length - 1];
        //   BodyMonthly.innerHTML = "";
        DataAD.forEach(e => {
            let filIt = item.filter(z => e.id == z.id);
            let mmm = 0, chhhh = 0, woooom = 0, inTopic = 0, outTopic = 0;
            for (let i of filIt) {
                let sc = i.seminarCount || 1;
                mmm += +(i.men);
                chhhh += +(i.child);
                woooom += +(i.women);
                inTopic += i.in ? sc : 0;
                outTopic += i.out ? sc : 0;
            }

            // console.log(mmm, chhhh, woooom, inTopic, outTopic);
            BodyMonthly.innerHTML += `
                <tr class = "${(inTopic || outTopic) ? "intopicandout" : ""}">
                    <td style="width: 15%;">${e.MainTopic}</td>
                    <td style="width: 15%;">${e.Subtopic}</td>
                    <td style="width: 6%;">${inTopic ? inTopic : ""}</td>
                    <td style="width: 6%;"></td>
                    <td style="width: 6%;">${outTopic ? outTopic : ""}</td>
                    <td style="width: 6%;"></td>
                    <td style="width: 6%;"></td>
                    <td style="width: 6%;">${(chhhh || mmm || woooom) ? chhhh : ""}</td>
                    <td style="width: 6%;">${(chhhh || mmm || woooom) ? mmm : ""}</td>
                    <td style="width: 6%;">${(chhhh || mmm || woooom) ? woooom : ""}</td>
                    <td style="width: 6%;"></td>
                    <td style="width: 6%;"></td>
                    <td style="width: 10%;"></td>
                </tr>`;
        })
        // console.log("==========");
        monthly.innerHTML += `<hr/>`
    })

    //////////////////////////////////////////////////////////////
    var yearly = document.querySelector(".yearly");
    yearly.innerHTML = "";
    yearly.innerHTML += `
    <p class="mt-2 text-center fw-bold">الخطة مجتمعة في جدول واحد</p>
    `;
    yearly.innerHTML += `
    <table class="table-group-divider" value = "الخطة المجمعة">
    <thead>
      <tr>
        <th style="width: 15%;">الموضوع الرئيسي</th>
        <th style="width: 15%;">الموضوع الفرعي</th>
        <th style="width: 6%;">ندوة داخلية</th>
        <th style="width: 6%;">مشورة عامة</th>
        <th style="width: 6%;">ندوة خارجية</th>
        <th style="width: 6%;">لفاء جماهيري</th>
        <th style="width: 6%;">ندوات قوافل</th>
        <th style="width: 6%;">أطفال</th>
        <th style="width: 6%;">ذكور</th>
        <th style="width: 6%;">إناث</th>
        <th style="width: 6%;">استخدام وسائل إعلامية</th>
        <th style="width: 6%;">انشطة إعلامية رقمية</th>
        <th style="width: 10%;">ملاحظــات</th>
      </tr>
    </thead>
    <tbody>
    </tbody>
  </table>
  `;
    var boold = document.querySelector(".yearly tbody");
    DataAD.forEach(element => {
        var manscount = 0, childcount = 0, womencount = 0, intpcoun = 0, outtpcoun = 0;
        for (let i of DoneAll) {
            if (i.id == element.id) {
                let sc = i.seminarCount || 1;
                manscount += +(i.men);
                childcount += +(i.child);
                womencount += +(i.women);
                intpcoun += i.in ? sc : 0;
                outtpcoun += i.out ? sc : 0;
            }
        }
        boold.innerHTML += `
        <tr class = "${(intpcoun || outtpcoun) ? "intopicandout" : ""}">
            <td style="width: 15%;">${element.MainTopic}</td>
            <td style="width: 15%;">${element.Subtopic}</td>
            <td style="width: 6%;">${intpcoun ? intpcoun : ""}</td>
            <td style="width: 6%;"></td>
            <td style="width: 6%;">${outtpcoun ? outtpcoun : ""}</td>
            <td style="width: 6%;"></td>
            <td style="width: 6%;"></td>
            <td style="width: 6%;">${(childcount || manscount || womencount) ? childcount : ""}</td>
            <td style="width: 6%;">${(childcount || manscount || womencount) ? manscount : ""}</td>
            <td style="width: 6%;">${(childcount || manscount || womencount) ? womencount : ""}</td>
            <td style="width: 6%;"></td>
            <td style="width: 6%;"></td>
            <td style="width: 10%;"></td>
        </tr>`;
    })

    document.getElementById("download").classList.remove("d-none")
})

function downloadTablesOfTopics() {
    let tables = document.querySelectorAll(".monthlyCon table")
    // tables.add(tobo);

    var workbook = XLSX.utils.book_new();

    tables.forEach((table, index) => {
        var worksheet = XLSX.utils.table_to_sheet(table);

        worksheet['!margins'] = { RTL: true };

        for (let cell in worksheet) {
            if (worksheet.hasOwnProperty(cell) && cell[0] !== '!') {
                worksheet[cell].s = {
                    alignment: {
                        vertical: 'center',
                        horizontal: 'center'
                    }
                };
            }
        }

        let range = XLSX.utils.decode_range(worksheet['!ref']);
        worksheet['!cols'] = [];
        for (let C = range.s.c; C <= range.e.c; ++C) {
            let maxWidth = 10;
            for (let R = range.s.r; R <= range.e.r; ++R) {
                let cell_address = { c: C, r: R };
                let cell_ref = XLSX.utils.encode_cell(cell_address);
                let cell = worksheet[cell_ref];
                if (cell && cell.v) {
                    let cellValue = cell.v.toString();
                    maxWidth = Math.max(maxWidth, cellValue.length);
                }
            }
            worksheet['!cols'][C] = { width: maxWidth };
        }

        XLSX.utils.book_append_sheet(workbook, worksheet, table.getAttribute("value"));
    });

    workbook.Workbook = {
        Views: [{ RTL: true }]
    };

    var excelFile = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });

    var blob = new Blob([s2ab(excelFile)], { type: "application/octet-stream" });

    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = "الخطة.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);

}

function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}

function downloadTableOfDetails() {
    var workbook = XLSX.utils.book_new();
    var worksheet = XLSX.utils.table_to_sheet(tobo);
    worksheet['!margins'] = { RTL: true };
    for (let cell in worksheet) {
        if (worksheet.hasOwnProperty(cell) && cell[0] !== '!') {
            worksheet[cell].s = {
                alignment: {
                    vertical: 'center',
                    horizontal: 'center'
                }
            };
        }
    }
    let range = XLSX.utils.decode_range(worksheet['!ref']);
    worksheet['!cols'] = [];
    for (let C = range.s.c; C <= range.e.c; ++C) {
        let maxWidth = 10;
        for (let R = range.s.r; R <= range.e.r; ++R) {
            let cell_address = { c: C, r: R };
            let cell_ref = XLSX.utils.encode_cell(cell_address);
            let cell = worksheet[cell_ref];
            if (cell && cell.v) {
                let cellValue = cell.v.toString();
                maxWidth = Math.max(maxWidth, cellValue.length);
            }
        }
        worksheet['!cols'][C] = { width: maxWidth };
    }
    XLSX.utils.book_append_sheet(workbook, worksheet, "الموضوعات اليومية");
    workbook.Workbook = {
        Views: [{ RTL: true }]
    };
    var excelFile = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });
    var blob = new Blob([s2ab(excelFile)], { type: "application/octet-stream" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = "الموضوعات.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

