# Basic Excel Import For OpenCart v3.0

## INSTALLING

**OC MOD INSTALLING**

When you download files, you will find "excel_import.ocmod.zip" file.

* Login you administration panel,
* Go to **Extentions -> Installer**
* Select "excel_import.ocmod.zip" file and upload.
* After installing module go to **System -> Users -> User Groups -> Edit User Group**
* Find "catalog/excel" both Access Permisison and Modify Permisison lists and check it.
* It's Done !

**Standart Ftp Uploading**

**NOTE :** If you don't use oc_mod install method, you should add a link manually to navigation for access to module. Because there is no modification for edit menu items.

**Link :** index.php?route=catalog/excel&user_token=*******

* Download file from repository
* Find "upload" folder and upload inside files to your opencart root dir.
* After upload files, go to **System -> Users -> User Groups -> Edit User Group**
* Find "catalog/excel" both Access Permisison and Modify Permisison lists and check it.
* It's Done !


## USAGE

* First we must set some settings before upload an excel file,
* Choose which product form data assign to in your excel,
* After edit all rows or just using in excel, save settings,
* Select your upload file and submit form.

## USAGE NOTES

**CATEGORIES**

You can use only one row for a product, so if you want to set product's category information from excel read above.

**EX. product info with product categories**

---
ROW A : Product Model
ROW B : Product Name
ROW C : Main Category Name
ROW D : Sub Category Name
ROW E : Sub Category Name
---

In this example our products goes to; **Main Category -> Sub Category -> Sub Category**, but if you whant to **set multiple main category** take a look example above.


---
ROW_A : Product Model
ROW_B : Product Name
ROW_C : Main Cat|Main Cat 2
ROW_D : Sub Cat 1|Sub Cat 2
ROW_E : Sub Cat 1,2|Sub Cat 2,2
---

In this example our product goes to; **Main Cat -> Sub Cat 1 -> Sub Cat 1,2** and **Main Cat 2 -> Sub Cat 2 -> Sub Cat 2,2** categories.

Sortly for multiple category you can **use sperator** " | " between category names.


**Product Options**

For add an option to product, you should add new row for product option. If another row has same product model, system get this row as product option.

**Example for product option**
---
ROW_A (Product Model) : MODEL0001 : MODEL0001
ROW_B : Product Name : OPTION NAME|OPTION VALUE NAME
ROW_C : 127.00 : +50.00
ROW_D : 10 : 10
---

If you want to add more options you can add more row What has to same model number.

* You should use same quantitiy columb for options quantitiy info.
* You should use same price columb for options price info.
* If you want to " - " prefix for price, you can set price columb "-50.00"
* Dont use curreny sym in price columbs.








