const fs = require("fs");
const xml2js = require("xml2js");
const path = require("path");
const { NtlmClient } = require("axios-ntlm");

const client = NtlmClient({
  username: "", // SharePoint username
  password: "", // SharePoint password
  domain: "", // Domain
});

// .iqy dosyasından URL'yi al
const readIqyFile = (iqyFilePath) => {
  const iqyContent = fs.readFileSync(iqyFilePath, "utf-8");
  const url = iqyContent.split("\n")[2].trim(); // URL genellikle üçüncü satırdadır
  return new URL(url);
};

// owssvr.xml dosyasını indir ve dosya listesini al
const getFileListFromOwssvr = async (url) => {
  try {
    const response = await client.get(url, {
      headers: {
        Accept: "application/json;odata=verbose",
      },
    });

    const xmlData = response.data;

    // XML'i JSON formatına çevir
    const parser = new xml2js.Parser();
    const jsonData = await parser.parseStringPromise(xmlData);

    // SharePoint'ten gelen dosya listesi genellikle `<z:row>` tagleri altında olur
    const rows = jsonData["xml"]["rs:data"][0]["z:row"];

    const fileList = rows.map((row) => {
      return {
        name: row["$"]["ows_LinkFilename"],
        subDir: row["$"]["ows_ContentType"],
      };
    });

    return fileList;
  } catch (error) {
    console.error("Error fetching owssvr.xml:", error);
  }
};

// SharePoint'teki her bir dosyayı indir
const downloadFile = async (baseUrl, fileName, downloadDir, subDir) => {
  // dizin yoksa oluştur
  if (!fs.existsSync(path.join(downloadDir, subDir))) {
    fs.mkdirSync(path.join(downloadDir, subDir), { recursive: true });
  }

  const fileUrl = `${baseUrl}/${fileName}`;
  const filePath = path.join(downloadDir, subDir, fileName);

  try {
    const response = await client.get(fileUrl, {
      responseType: "arraybuffer",
    });

    // Dosyayı kaydet
    fs.writeFileSync(filePath, response.data);
    console.log(`Downloaded: ${fileName}`);
  } catch (error) {
    console.error(`Error downloading file (${fileName}):`, error);
  }
};

// Ana iş akışı
const main = async () => {
  const iqyFilePath = "./query.iqy"; // .iqy dosyasının yolu
  const downloadDir = "./downloads"; // Dosyaların indirileceği klasör

  // .iqy dosyasından URL'yi al
  const owssvrUrl = readIqyFile(iqyFilePath);

  // URL'nin baz kısmını ayır
  const baseUrl = owssvrUrl.origin + "/Dokuman/Dokumanlar";

  console.log(baseUrl);

  // Dosya listesini al
  const fileList = await getFileListFromOwssvr(owssvrUrl);

  // Dosya listesini indir
  for (const fileDetail of fileList) {
    await downloadFile(
      baseUrl,
      fileDetail.name,
      downloadDir,
      fileDetail.subDir
    );
  }
};

main();
