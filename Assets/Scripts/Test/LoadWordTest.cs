using UnityEngine;
using Aspose.Words;
using Aspose.Words.Saving;
using System.IO;
using UnityEngine.UI;

/// <summary>
/// 加载word
/// </summary>
public class LoadWordTest : MonoBehaviour
{
    public static LoadWordTest Instance;

    public GameObject imageprefab;
    public Transform content;

    private void Start()
    {
        Instance = this;
    }

    /// <summary>
    /// 加载Word
    /// </summary>
    public void LoadWordGo(string path)
    {
        //清理content下的旧物体
        for (int i = 0; i < content.childCount; i++)
        {
            Destroy(content.GetChild(i).gameObject);
        }


        Document doc = new Document(path);

        ImageSaveOptions iso = new ImageSaveOptions(SaveFormat.Jpeg);
        iso.PrettyFormat = true; //漂亮的格式
        iso.UseAntiAliasing = true; //抗锯齿
        for (int i = 0; i < doc.PageCount; i++)
        {
            iso.PageIndex = i;
            //声明内存流，将图片转换为内存流，再由流转换为byte数组，然后用texture2d加载byte数组
            using (MemoryStream ms = new MemoryStream())
            {
                doc.Save(ms, iso);
                byte[] buff = new byte[ms.Length];
                ms.Seek(0, SeekOrigin.Begin);
                ms.Read(buff, 0, (int)ms.Length);

                //注意这个iamge的命名空间为system.drawing不是unity.ui,这个图片的目的是提供图片的宽高
                System.Drawing.Image sizeImage = System.Drawing.Image.FromStream(ms);

                Texture2D texture2D = new Texture2D(sizeImage.Width, sizeImage.Height);
                texture2D.LoadImage(buff);

                Image image = Instantiate(imageprefab, content).GetComponent<Image>();
                //根据转化出来的图片的大小设置unity image的大小
                image.rectTransform.sizeDelta = new Vector2(sizeImage.Width, sizeImage.Height);
                //用texture2d为精灵赋值
                image.sprite = Sprite.Create(texture2D, new Rect(0, 0, texture2D.width, texture2D.height), Vector2.zero);
            }
        }

    }


}
