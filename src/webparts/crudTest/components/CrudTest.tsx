import * as React from 'react';

import type { ICrudTestProps } from '../interfaces/ICrudTestProps';
import type { IProduct, IProductRes } from '../interfaces/ProductInterfaces';

import styles from './CrudTest.module.scss';

import { getSP, SPFI} from '../../../pnpjsConfig';

import { DefaultButton, DetailsList, DetailsListLayoutMode, Dialog, DialogFooter, DialogType, Dropdown, IColumn, IconButton, Label, PrimaryButton, SelectionMode, TextField } from '@fluentui/react';
import { initializeIcons } from '@fluentui/font-icons-mdl2';


const CrudTest = (props:ICrudTestProps): React.ReactElement => {
  
    initializeIcons();
    const sp:SPFI = getSP(props.spcontext);

    const [reload, setReload] = React.useState<boolean>(false);
    const [products, setProducts] = React.useState<Array<IProduct>>([]); 
    const [newProductName, setNewProductName] = React.useState<string>('');
    const [newDescription, setNewDescription] = React.useState<string>('');
    const [newCategory, setNewCategory] = React.useState<string>('');
    const [newQuantity, setNewQuantity] = React.useState<number>(0);
    const [isAddDialogHidden, setIsAddDialogHidden] = React.useState<boolean>(true);
    const [currentId, setCurrentId] = React.useState<number>();
    const [isEditDialogHidden, setIsEditDialogHidden] = React.useState<boolean>(true);
    const [editedProductName, setEditedProductName] = React.useState<string>('');
    const [editedDescription, setEditedDescription] = React.useState<string>('');
    const [editedCategory, setEditedCategory] = React.useState<string>('');
    const [editedQuantity, setEditedQuantity] = React.useState<number>(0);
    const [isDeleteDialogHidden, setIsDeleteDialogHidden] = React.useState<boolean>(true);
    const [deleteProductName, setDeleteProductName] = React.useState<string>('');
    const [productDetail, setProductDeatail] = React.useState<IProduct>();
    const [IsDetailDialogHidden, setIsDetailDialogHidden] = React.useState<boolean>(true);
    
    const [productNameError, setProductNameError] = React.useState<boolean>(false);
    const [quantityError, setQuantityError] = React.useState<boolean>(false);
    const [categoryError, setCategoryError] = React.useState<boolean>(false);
    

    const columns: IColumn[] = []
    columns.push(
      { key: 'column1', name: 'Product Name', fieldName: 'productName', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column2', name: 'Quantity', fieldName: 'quantity', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column3', name: 'Date Of Update', fieldName: 'dateOfUpdate', minWidth: 150, maxWidth: 200, isResizable: true },
      { key: 'column4', name: 'Category', fieldName: 'category', minWidth: 100, maxWidth: 200, isResizable: true },
      { 
        key: 'column5',
      name: '',
      minWidth: 120,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: IProduct) => (
        <>
        <IconButton
          iconProps={{ iconName: 'Edit' }}
          title="Edit"
          ariaLabel="Edit"
          onClick={() => openEditDialog(item)}
        />
        <IconButton
          iconProps={{ iconName: 'Delete' }}
          title="Delete"
          ariaLabel="Delete"
          onClick={() => openDeleteDialog(item)}
          className={styles.deleteButton}
        />
        <IconButton
          iconProps={{ iconName: 'Info' }}
          title="More Details"
          ariaLabel="More Details"
          onClick={() => openDetailDialog(item)}
          className={styles.deleteButton}
        />
        </>
      ),
      }
    );
    
//CRUD SECTION

//GET ALL PRODUCTS FROM LIST
  const getAllProductsFromList = async() =>{
    try{
      const productsList = await sp.web.lists.getByTitle("Products").items.select("Id", "Title", "Description", "Quantity", "Category", "DateOfUpdata") ();

      setProducts(productsList.map((each: IProductRes) => ({
        id: each.Id,
        productName: each.Title,
        description: each.Description,
        quantity: each.Quantity,  
        category: each.Category,
        dateOfUpdate: formatDateTime(each.DateOfUpdata),
      })));

    }catch(e){
      console.log(e)
    }finally{
      console.log("List Products fetched", products);
    }
  }


//ADD NEW PRODUCT
  const addNewProductToList = async() => {

    if (!validateFieldsAddDialog()) {
      console.log("Required fields in add dialog are empty")
      return;
    }

    const list = sp.web.lists.getByTitle("Products");
    try{
      await list.items.add({
        Title: newProductName,
        Description: newDescription,
        Quantity: newQuantity,
        Category: newCategory,
        DateOfUpdata: new Date()
      })
      setReload(!reload);
    } catch(e){
      console.log(e);
    }finally{
      console.log("Product added successfully")
      setNewCategory('')
      setNewDescription('')
      setNewProductName('')
      setNewQuantity(0)
      setIsAddDialogHidden(true);
    }
  }

//EDIT PRODUCT FROM LIST
  const editProductList = async () =>{

    if (!validateFieldsEditDialog()) {
      console.log("Required fields in edit dialog are empty")
      return;
    }

    const list =  sp.web.lists.getByTitle("Products");
    if (currentId !== undefined && currentId !== null) {
      try{
        await list.items.getById(currentId).update({
          Title: editedProductName,
          Description: editedDescription,
          Quantity: editedQuantity,
          Category: editedCategory,
          DateOfUpdata: new Date(),
        })
        
      }catch(e){
        console.log(e)
      }finally{
        console.log('Product ' + currentId + ' is successfully edited');
        setReload(!reload);
        setIsEditDialogHidden(true);
        setEditedProductName('');
        setEditedDescription('');
        setEditedQuantity(0);
        setEditedCategory('');

      }
    }
  }
//DELETE PRODUCT FROM LIST
  const deleteProductList = async() =>{
    const list = sp.web.lists.getByTitle("Products");
    if (currentId !== undefined && currentId !== null) {
      try{
        await list.items.getById(currentId).delete();
        
      }catch(e){
        console.log(e)
      }finally{
        console.log('Product is successfully deleted');
        setReload(!reload);
        setIsDeleteDialogHidden(true);
      }
    }
  }


//FORMAT DATE
function formatDateTime(dateTimeString: string) {
  let formattedString = dateTimeString.replace("T", " ");
  formattedString = formattedString.replace("Z", "");
  const lastColonIndex = formattedString.lastIndexOf(":");
  if (lastColonIndex !== -1) {
      formattedString = formattedString.slice(0, lastColonIndex);
  }
  return formattedString.trim(); 
}

//CHECKING REQUIRED FIELDS
//FIELDS FROM ADD DIALOG
  function validateProductNameAddDialog():boolean {
    if (!newProductName.trim()) {
      setProductNameError(true);
      return false;
    } else {
      setProductNameError(false);
      return true;
    }
  }

  function validateQuantityAddDialog():boolean {
    if (newQuantity <= 0 || isNaN(newQuantity)) {
      setQuantityError(true);
      return false;
    } else {
      setQuantityError(false);
      return true;
    }

  }

  function validateCategoryAddDialog():boolean{
    if (!newCategory.trim()) {
      setCategoryError(true);
      return false;
    } else {
      setCategoryError(false);
      return true;
    }
  }

  function validateFieldsAddDialog():boolean{
    let isValid = true;

    if (!validateProductNameAddDialog()){
      isValid = validateProductNameAddDialog();
    }

    if (!validateQuantityAddDialog()) {
      isValid = validateQuantityAddDialog();
    }

    if (!validateCategoryAddDialog()) {
      isValid = validateCategoryAddDialog();
    }

    return isValid;
  }

//FIELDS FROM EDIT DIALOG
  function validateProductNameEditDialog():boolean {
    if (!editedProductName.trim()) {
      setProductNameError(true);
      return false;
    } else {
      setProductNameError(false);
      return true;
    }
  }

  function validateQuantityEditDialog():boolean {
    if (editedQuantity <= 0 || isNaN(editedQuantity)) {
      setQuantityError(true);
      return false;
    } else {
      setQuantityError(false);
      return true;
    }

  }

  function validateCategoryEditDialog():boolean{
    if (!editedCategory.trim()) {
      setCategoryError(true);
      return false;
    } else {
      setCategoryError(false);
      return true;
    }
  }

  function validateFieldsEditDialog():boolean{
    let isValid = true;

    if (!validateProductNameEditDialog()){
      isValid = validateProductNameEditDialog();
    }

    if (!validateQuantityEditDialog()) {
      isValid = validateQuantityEditDialog();
    }

    if (!validateCategoryEditDialog()) {
      isValid = validateCategoryEditDialog();
    }

    return isValid;
  }

// open/close dialog functions
function openAddDialog(): void {
  setIsAddDialogHidden(false);
}

function closeAddDialog(): void {
  setIsAddDialogHidden(true);
  setNewCategory('');
  setNewDescription('');
  setNewProductName('');
  setNewQuantity(0);
  setProductNameError(false);
  setCategoryError(false);
  setQuantityError(false);
}

function openEditDialog(item: IProduct): void{
  setCurrentId(item.id);
  setIsEditDialogHidden(false); 
  const product: IProduct | undefined = products.find((each) => each.id === item.id); 
  if (product) {
    setEditedProductName(product.productName);
    setEditedDescription(product.description);
    setEditedQuantity(product.quantity);
    setEditedCategory(product.category);
  }
}

  function closeEditDialog(): void{
    setIsEditDialogHidden(true);
    setEditedProductName('');
    setEditedDescription('');
    setEditedQuantity(0);
    setEditedCategory('');
    setProductNameError(false);
    setCategoryError(false);
    setQuantityError(false);
  }

function openDeleteDialog(item: IProduct): void{
  setCurrentId(item.id);
  setDeleteProductName(item.productName);
  setIsDeleteDialogHidden(false);
}

function openDetailDialog(item: IProduct): void{
  setProductDeatail(item);
  setIsDetailDialogHidden(false);
}


//HANDLERS
function handelNewProductName(event: React.ChangeEvent<HTMLInputElement>): void{
  setNewProductName(event.target.value);
}

function handelNewDescription(event: React.ChangeEvent<HTMLInputElement>): void{
  setNewDescription(event.target.value);
}

function handelNewQuantiry(event: React.ChangeEvent<HTMLInputElement>): void{
  setNewQuantity(parseInt(event.target.value));
}

function handelNewCategory(event: React.FormEvent<HTMLDivElement>, option?: { key: string | number, text: string }): void{
  if (option) {
    setNewCategory(option.key as string);
  }
}

function handelEditedProductName(event: React.ChangeEvent<HTMLInputElement>): void{
  setEditedProductName(event.target.value);
}

function handelEditedDescription(event: React.ChangeEvent<HTMLInputElement>): void{
  setEditedDescription(event.target.value);
}

function handelEditedQuantiry(event: React.ChangeEvent<HTMLInputElement>): void{
  setEditedQuantity(parseInt(event.target.value));
}

function handelEditedCategory(event: React.FormEvent<HTMLDivElement>, option?: { key: string | number, text: string }): void{
  if (option) {
    setEditedCategory(option.key as string);
  }
}

//useEffects
React.useEffect(() =>{
  getAllProductsFromList();
},[reload]);



    return (
      <div>
      <DetailsList
      items={products}
      columns={columns}
      layoutMode={DetailsListLayoutMode.justified}
      selectionMode={SelectionMode.none} 
      />
      <IconButton
        iconProps={{iconName: 'Add'}}
        title='Add new product'
        ariaLabel='Add new product'
        onClick={()=>openAddDialog()}
        styles={{ root: { margin: 20, backgroundColor: 'lightgreen', color: 'white' } }}
      />
      <div>
        <Dialog
          hidden={isAddDialogHidden}
          onDismiss={() => setIsAddDialogHidden(true)}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Add New Product',
          }}
        >
          <div>
            <TextField
              className={`${productNameError ? styles.errorField : styles.textField} `}
              required
              placeholder='Product Name'
              value= {newProductName}
              onChange= {handelNewProductName}
              validateOnFocusIn = {productNameError} 
              onBlur={validateProductNameAddDialog}
            />
            <TextField
              className={`${styles.textField}`}
              placeholder='Description'
              multiline
              rows={4}
              value= {newDescription}
              onChange= {handelNewDescription}
            />
            <TextField
              className={`${quantityError ? styles.errorField : styles.textField}`}
              required
              placeholder='Quantity'
              type='number'
              min={0}
              value= {newQuantity.toString()}
              onChange= {handelNewQuantiry} 
              onBlur={validateQuantityAddDialog}
            />
            <Dropdown
              required
              className={`${categoryError ? styles.errorField : styles.dropdown}`}
              options={props.choices}
              placeholder='Category'
              onChange={handelNewCategory}
              selectedKey={newCategory}   
              onBlur={validateCategoryAddDialog}
            />
          </div>
          <DialogFooter>
              <PrimaryButton text="Add" onClick={() => addNewProductToList()} />
              <DefaultButton text="Cancel" onClick={() => closeAddDialog()} />
            </DialogFooter>
        </Dialog>
      </div>
      <div>
        <Dialog
          hidden={isEditDialogHidden}
          onDismiss={() => setIsEditDialogHidden(true)}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Add New Product',
          }}
        >
          <div>
            <TextField
              required
              className= {`${productNameError ? styles.errorField : styles.textField}`}
              placeholder='Product Name'
              value= {editedProductName}
              onChange= {handelEditedProductName}
              onBlur={validateProductNameEditDialog}
            />
            <TextField
              className={`${styles.textField}`}
              placeholder='Description'
              multiline
              rows={4}
              value= {editedDescription}
              onChange= {handelEditedDescription}
            />
            <TextField
              required
              className= {`${quantityError ? styles.errorField : styles.textField}`}
              placeholder='Quantity'
              type='number'
              min={0}
              value= {editedQuantity.toString()}
              onChange= {handelEditedQuantiry}
              onBlur={validateQuantityEditDialog}
            />
            <Dropdown
              required
              className= {`${categoryError ? styles.errorField : styles.dropdown}`}
              options={props.choices}
              placeholder='Category'
              onChange={handelEditedCategory}
              selectedKey={editedCategory}
              onBlur={validateCategoryEditDialog}
            />
          </div>
          <DialogFooter>
            <PrimaryButton text="Edit" onClick={() => editProductList()} />
            <DefaultButton text="Cancel" onClick={() => closeEditDialog()} />
          </DialogFooter>
        </Dialog>
      </div>
      <div className='centered-div' >
        <Dialog
          
          hidden={isDeleteDialogHidden}
          onDismiss={() => setIsDeleteDialogHidden(true)}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Delete Product',
            className: styles.textAlignCenter,
          }}
        >
          <Label className={styles.textAlignCenter}>Are you sure that you want delete this product?</Label>
          <Label className={styles.textAlignCenter}><h1>{deleteProductName}</h1></Label>
          <DialogFooter
            styles={{
              actionsRight: {
                justifyContent: 'center', 
                display: 'flex',
                width: '100%', 
              },
            }}
          >
            <PrimaryButton text="Delete" onClick={() => deleteProductList()} />
            <DefaultButton text="Cancel" onClick={() => setIsDeleteDialogHidden(true)} />
          </DialogFooter>
        </Dialog>
      </div>
      <div>
        <Dialog
          hidden={IsDetailDialogHidden}
          onDismiss={() => setIsDetailDialogHidden(true)}
          dialogContentProps={{
            type: DialogType.normal,
            title: productDetail?.productName,
          }}
        >
          <div className={styles.customDiv}>
            <Label className={styles.textAlignCenter}>{productDetail?.description}</Label>
            <div className= {styles.textAlignRight}>
              <h6>{productDetail?.dateOfUpdate}</h6>
            </div>
          </div>
          <DialogFooter
            styles={{
              actionsRight: {
                justifyContent: 'center', 
                display: 'flex',
                width: '100%', 
              },
            }}
          >
            <PrimaryButton text="Close" onClick={() => setIsDetailDialogHidden(true)} />
          </DialogFooter>
        </Dialog>
      </div>
    </div>
    ); 
}

export default CrudTest