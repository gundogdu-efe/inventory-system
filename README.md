**Envanter Yönetim Sistemi / Inventory Management System**

[TR]

Bu proje, stok ve envanter takibini kolaylaştırmak amacıyla geliştirilmiş basit bir **Envanter Yönetim Sistemi** uygulamasıdır. Uygulama Türkçe'dir ve ASP.NET MVC (.NET Framework) mimarisiyle geliştirilmiş olup Microsoft SQL Server veritabanı ile Bootstrap kütüphanesi kullanmaktadır.

[EN]

This project is a simple **Inventory Management System** application developed to facilitate stock and inventory tracking. The application is in Turkish and has been developed with ASP.NET MVC (.NET Framework) architecture and uses Microsoft SQL Server database and Bootstrap library.

**Kullanılan Teknolojiler / Technologies Used**

- ASP.NET MVC (.NET Framework)
- C#
- Microsoft SQL Server
- Bootstrap
- HTML / CSS

**Özellikler / Features**

**1) Giriş Sayfası / Login Page**

[TR]

Kayıt olmuş kullanıcıların giriş yaptığı sayfadır.

[EN]

The page where registered users log in

![girişyap](https://github.com/user-attachments/assets/cd792659-5d36-4c31-b605-4bb6a1ec705e)

**2) Kayıt Sayfası / Register Page**

[TR]

Kullanıcıların kayıt olduğu sayfadır.

[EN]

This is the page where users register.

![kayıtol](https://github.com/user-attachments/assets/6d0fb67b-4f22-4f1b-abaf-4c36f26312f3)

**3) Şifremi Unuttum / Forgot Password Page**

[TR]

Şifresini unutan kullanıcıların önce maillerini kontrol ettiği sonrasında ise şifrelerini değiştirdiği sayfadır.

[EN]

This is the page where users who have forgotten their password. First, check their emails and then change their passwords.

![şifremiunuttum](https://github.com/user-attachments/assets/52ffd335-13c2-4bfe-a986-fb0082eb0ed6)

![şifremiunuttum2](https://github.com/user-attachments/assets/495e8d67-66f7-4b77-8fdd-b8ce69eec6bd)

**4) Ana Sayfa / Home Page**

[TR]

Ana sayfada kullanıcılar direkt olarak envanter listesini görebiliyor.

[EN]

On the Home page, users can see the inventory list directly.

![Anasayfa](https://github.com/user-attachments/assets/4db699e5-954c-4941-8497-60e6fcc07223)

**5) Envanter Listesi / Inventory List**

[TR]

Kullanıcılar Envanter Listesi bölümünde envanterleri liste halinde görüntüleyebilir ve envanter listesinde güncellemeler yapabilir. Güncelleme yapabileceği değişiklikler: Birimler, Kategoriler, Türler ve Açıklama bölümüdür.

[EN]

In the inventory list section, users can view the inventories in a list and make updates to the inventory list. The changes they can update are: Units, Categories, Types, and the Description section.

![envanterlistesi](https://github.com/user-attachments/assets/31c7064a-f082-49d2-bb5e-72d6b2d1186a)

![envantergüncelle1](https://github.com/user-attachments/assets/0046984c-e6e4-4635-8aad-6bdf99a77dd3)

![envantergüncelle2](https://github.com/user-attachments/assets/a115b0c1-8a6d-4f13-860e-ba1e734a53a3)

**6) Envanter Ekle / Add Inventory**

[TR]

Kullanıcıların envantere, Kategori, Tür ID, Kodu, İlgili Birim ve Açıklamasıyle envantere ekleme yaptığı bölümdür.

[EN] 

This is a section where users add items to the inventory with their category, type id, inventory code, relevant unit, and description.

![envanterekle](https://github.com/user-attachments/assets/3b329aea-9343-44a8-ae63-73e532a83622)

![envanterekle2](https://github.com/user-attachments/assets/560e1e84-2cfe-48aa-939b-c097abbccd82)

**7) Zimmet Listesi / Inventory Assignment List**

[TR]

Envanterde ekleme ya da değişiklik yapan kişilerin takip edildiği ve güncellemeler yapılabildiği bir bölümdür.

[EN]

It is a section where people who make additions or changes to the inventory are tracked, and updates can be made.


![zimmetlistesi](https://github.com/user-attachments/assets/6b8f6307-3b1b-4c81-abab-f8fd337ac47b)

![zimmetdeğiştir1](https://github.com/user-attachments/assets/e4501cf7-d3ff-4602-81c5-98f654f074ad)

![zimmetdeğiştir2](https://github.com/user-attachments/assets/c65f4979-0171-42eb-b1c4-05b04805cf2c)

![zimmetdeğiştir3](https://github.com/user-attachments/assets/b29dbf84-d23e-40bc-a179-f777c9cbecc4)

**8) Zimmet Ekleme / Adding Inventory Assignment**

[TR]

Zimmet Türü, Birim, Personel, Lokasyon, Oda Numarası, Açıklama ve Tarih bilgileriyle bir Zimmet Eklemesi yapılan bölümdür.

[EN]

This is the section where an Inventory Assignment is added with the information of Inventory Assignment Type, Unit, Personnel, Location, Room Number, Description, and Date.

![zimmetekle](https://github.com/user-attachments/assets/29a12e97-f4c1-41b7-9768-2421419a27ba)

**9) Zimmet Baskı / Inventory Assignment Print**

[TR]

Bu bölümde önceden oluşturulmuş olan listeler bir Excel dosyasına yazdırılır ve hazır bir şablon olarak baskısı alınır.

[EN]

The lists created in this section are printed to an Excel file and printed as a ready report.

![zimmetbaskı](https://github.com/user-attachments/assets/f1753ff6-933b-41fa-b1af-efe742e39fa9)

![baskıform](https://github.com/user-attachments/assets/6de4eb2c-b952-4474-b432-b7d32e8f0f8f)

**10) Marka Model İşlemleri / Brand Model Operations**

[TR]

Envanterdeki ürünlerin marka ve model takibinin yapıldığı ve yeni marka ve modellerin eklendiği bölümdür.

[EN]

This is the section where the brands and models of the products in the inventory are tracked and new brands and models are added.

![markamodel](https://github.com/user-attachments/assets/9b35352b-01f0-4abe-87e1-47251e33c234)

![markamodelekle](https://github.com/user-attachments/assets/e7cb357b-c105-4802-934c-99ba56b0e7de)










































