
// 转换
function btnExcelToJson() {
    // 获取excel数据
    let excelData = (document.getElementById("input-excelData") as HTMLInputElement).value;
    // 获取转换后的json数据
    (document.getElementById("input-jsonData") as HTMLInputElement).value = excelToJson(excelData);
}

// 转换并复制
function btnExcelToJsonAndCopy() {
    btnExcelToJson();
    // 获取json数据
    let inputJsonData = document.getElementById("input-jsonData") as HTMLInputElement;
    // 选中并复制
    inputJsonData.select();
    document.execCommand("Copy");
}

/**
 * excel数据转json数据
 * excelToJson
 * @param excelData 
 * @returns 
 */
function excelToJson(excelData: string): string {

    // 将制表符替换为,
    let str = excelData.replace(new RegExp("\t", "g"), "/_/");
    // 将回车符替换为*_*
    str = str.replace(new RegExp("\n", "g"), "/_/*_*/_/");
    // console.log(str);

    // 转化为数组
    let arr = str.split("/_/");
    arr.pop();
    // console.log(arr);

    // 名称列表
    let names = [];
    // 类型列表
    let types = [];
    // 值列表
    let valueList: string[][] = [];

    // 获取名称列表
    let nameArr = JSON.parse(JSON.stringify(arr));
    for (let t of nameArr) {
        if (t == "*_*") {
            break;
        } else {
            names.push(t);
        }
    }

    // 获取类型列表
    let typeArr: string[] = JSON.parse(JSON.stringify(arr));
    typeArr.splice(0, names.length + 1);
    for (let t of typeArr) {
        if (t == "*_*") {
            break;
        } else {
            types.push(t);
        }
    }

    // 获取值列表
    let valueArr: string[] = JSON.parse(JSON.stringify(arr));
    valueArr.splice(0, names.length + 1 + types.length + 1);

    let tCount = 0;
    let isT = true;
    for (let t of valueArr) {

        if (isT) {
            valueList.push([]);
            isT = false;
            tCount++;
        }

        if (t == "*_*") {
            isT = true;
        } else {
            valueList[valueList.length - 1].push(t);
        }
    }

    // console.log(names);
    // console.log(types);
    // console.log(valueList);

    let jsonDatas = [];

    for (let i = 0; i < tCount; i++) {
        let jsonData = {};
        for (let j = 0; j < names.length; j++) {

            // 类型
            let type = types[j];
            // 数据
            let data: any = valueList[i][j];
            // 根据数据类型进行转换
            if (type == "string") {
                // data
            }
            if (type == "number") {
                data = Number(data);
            }
            if (type == "boolean") {
                if (data == "TRUE" || data == "true" || data == "1") {
                    data = true;
                } else {
                    if (data == "FALSE" || data == "false" || data == "0") {
                        data = false;
                    }
                }
            }
            if (type == "array") {
                if (data == "[]") {
                    data = [];
                } else {
                    data = data.replace("[", "");
                    data = data.replace("]", "");
                    data = data.split(",") as string[];

                    // 自动推断数组中的数据类型 字符串/数值/布尔值
                    for (let k = 0; k < data.length; k++) {
                        // number
                        if (!isNaN(Number(data[k]))) {
                            data[k] = Number(data[k]);
                        } else {
                            // boolean
                            if (data[k] == "TRUE" || data[k] == "true") {
                                data[k] = true;
                            } else {
                                if (data[k] == "FALSE" || data[k] == "false") {
                                    data[k] = false;
                                }
                            }
                        }
                    }
                }
            }
            if (type == "object") {
                // data
                data = JSON.parse(data);
            }

            (jsonData as any)[names[j]] = data;
        }

        jsonDatas.push(jsonData);
    }
    // console.log(jsonDatas);

    return JSON.stringify(jsonDatas);

}

