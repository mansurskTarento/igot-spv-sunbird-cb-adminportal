import { of, throwError } from 'rxjs'
import { BulkUploadOrgComponent } from './bulk-upload-org.component'

const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

describe('BulkUploadOrgComponent', () => {
  let component: BulkUploadOrgComponent
  let dialogRef: any
  let data: any
  let orgHieService: any
  let loaderService: any
  let snackbar: any

  const flush = () => new Promise(resolve => setTimeout(resolve, 0))

  beforeEach(() => {
    dialogRef = { close: jest.fn() }
    data = {
      bulkUploadConfig: {
        mainHeading: 'Upload',
        frameworkData: { identifier: 'root1', orgHierarchyFrameworkId: 'org1_fw' },
      },
    }
    orgHieService = {
      getBulkuploadProgress: jest.fn().mockReturnValue(
        of({ params: { status: 'SUCCESSFUL' }, result: { content: [{ fileName: 'a.xlsx' }] } })
      ),
      downloadSampleTemplate: jest.fn().mockReturnValue(of({ data: 'file' })),
      uploadFreameworkTemplate: jest.fn().mockReturnValue(of({ result: { fileName: 'a.xlsx' } })),
      downloadFileLog: jest.fn().mockReturnValue(of('blob-data')),
    }
    loaderService = { setLoaderState: jest.fn() }
    snackbar = { open: jest.fn() }
    jest.spyOn(console, 'log').mockImplementation(() => { })

    component = new BulkUploadOrgComponent(
      dialogRef as any,
      data as any,
      orgHieService as any,
      loaderService as any,
      snackbar as any
    )
  })

  afterEach(() => {
    jest.restoreAllMocks()
  })

  it('should create', () => {
    expect(component).toBeTruthy()
  })

  describe('ngOnInit', () => {
    it('should set config and load progress data', () => {
      component.ngOnInit()
      expect(component.bulkUploadConfig).toBe(data.bulkUploadConfig)
      expect(orgHieService.getBulkuploadProgress).toHaveBeenCalledWith('org1')
      expect(component.lastUploadList).toEqual([{ fileName: 'a.xlsx' }])
      expect(loaderService.setLoaderState).toHaveBeenCalledWith(false)
    })

    it('should default config to empty object when not provided', () => {
      component.data = {} as any
      const spy = jest.spyOn(component, 'getBulkuploadPrgressData').mockImplementation(() => { })
      component.ngOnInit()
      expect(component.bulkUploadConfig).toEqual({})
      expect(spy).toHaveBeenCalled()
    })
  })

  describe('getBulkuploadPrgressData', () => {
    beforeEach(() => {
      component.bulkUploadConfig = data.bulkUploadConfig
    })

    it('should set empty list for unsuccessful response', () => {
      orgHieService.getBulkuploadProgress.mockReturnValue(of({ params: { status: 'failed' } }))
      component.getBulkuploadPrgressData()
      expect(component.lastUploadList).toEqual([])
      expect(snackbar.open).toHaveBeenCalledWith('No progress data found')
    })

    it('should set empty list for null response', () => {
      orgHieService.getBulkuploadProgress.mockReturnValue(of(null))
      component.getBulkuploadPrgressData()
      expect(component.lastUploadList).toEqual([])
    })

    it('should handle response without status', () => {
      orgHieService.getBulkuploadProgress.mockReturnValue(of({ params: {} }))
      component.getBulkuploadPrgressData()
      expect(component.lastUploadList).toEqual([])
    })

    it('should handle error response', () => {
      orgHieService.getBulkuploadProgress.mockReturnValue(throwError({}))
      component.getBulkuploadPrgressData()
      expect(component.lastUploadList).toEqual([])
      expect(loaderService.setLoaderState).toHaveBeenCalledWith(false)
      expect(snackbar.open).toHaveBeenCalledWith('Error fetching progress data')
    })

    it('should pass empty string when org id is empty', () => {
      component.bulkUploadConfig = { frameworkData: { orgHierarchyFrameworkId: '' } }
      component.getBulkuploadPrgressData()
      expect(orgHieService.getBulkuploadProgress).toHaveBeenCalledWith('')
    })
  })

  describe('handleDownloadSampleFile', () => {
    it('should download sample file successfully', async () => {
      component.bulkUploadConfig = data.bulkUploadConfig
      await component.handleDownloadSampleFile()
      expect(loaderService.setLoaderState).toHaveBeenCalledWith(true)
      expect(orgHieService.downloadSampleTemplate).toHaveBeenCalledWith('root1')
      expect(snackbar.open).toHaveBeenCalledWith('Download successfully')
    })

    it('should show error message when download fails', async () => {
      component.bulkUploadConfig = data.bulkUploadConfig
      orgHieService.downloadSampleTemplate.mockReturnValue(throwError({ error: { params: { errMsg: 'bad' } } }))
      await component.handleDownloadSampleFile()
      expect(loaderService.setLoaderState).toHaveBeenCalledWith(false)
      expect(snackbar.open).toHaveBeenCalledWith('bad')
      expect(snackbar.open).not.toHaveBeenCalledWith('Download successfully')
    })

    it('should not show message when download fails without errMsg', async () => {
      component.bulkUploadConfig = data.bulkUploadConfig
      orgHieService.downloadSampleTemplate.mockReturnValue(throwError({}))
      await component.handleDownloadSampleFile()
      expect(snackbar.open).not.toHaveBeenCalled()
    })

    it('should do nothing when framework id is missing', async () => {
      component.bulkUploadConfig = {}
      await component.handleDownloadSampleFile()
      expect(orgHieService.downloadSampleTemplate).not.toHaveBeenCalled()
    })
  })

  it('handleFileClick should reset input value', () => {
    const event: any = { target: { value: 'C:/fake/a.xlsx' } }
    component.handleFileClick(event)
    expect(event.target.value).toBe('')
  })

  describe('onFileSelected', () => {
    it('should reject invalid file types', () => {
      const spy = jest.spyOn(component, 'uploadExcelFile')
      component.onFileSelected({ target: { files: [{ type: 'text/plain', size: 10 }] } })
      expect(spy).not.toHaveBeenCalled()
      expect(snackbar.open).toHaveBeenCalledWith('Please select a valid Excel file (.xlsx)', 'Close', { duration: 5000 })
    })

    it('should reject files larger than 5MB', () => {
      const spy = jest.spyOn(component, 'uploadExcelFile')
      component.onFileSelected({ target: { files: [{ type: XLSX_TYPE, size: 6 * 1024 * 1024 }] } })
      expect(spy).not.toHaveBeenCalled()
      expect(snackbar.open).toHaveBeenCalledWith('File size should not exceed 5MB', 'Close', { duration: 5000 })
    })

    it('should upload a valid file', () => {
      const spy = jest.spyOn(component, 'uploadExcelFile').mockResolvedValue(undefined as any)
      const file: any = { type: XLSX_TYPE, size: 10 }
      component.onFileSelected({ target: { files: [file] } })
      expect(spy).toHaveBeenCalledWith(file)
    })

    it('should do nothing when no file selected', () => {
      const spy = jest.spyOn(component, 'uploadExcelFile')
      component.onFileSelected({ target: { files: [] } })
      expect(spy).not.toHaveBeenCalled()
    })
  })

  describe('uploadExcelFile', () => {
    const file = new File(['x'], 'a.xlsx', { type: XLSX_TYPE })

    beforeEach(() => {
      component.bulkUploadConfig = data.bulkUploadConfig
    })

    it('should upload and refresh progress', async () => {
      const spy = jest.spyOn(component, 'getBulkuploadPrgressData').mockImplementation(() => { })
      await component.uploadExcelFile(file)
      expect(orgHieService.uploadFreameworkTemplate).toHaveBeenCalledWith(expect.any(FormData), data.bulkUploadConfig.frameworkData)
      expect(spy).toHaveBeenCalled()
      expect(snackbar.open).toHaveBeenCalledWith('File uploaded successfully. Please check after 5 minutes for the results.')
    })

    it('should show error message on upload failure', async () => {
      orgHieService.uploadFreameworkTemplate.mockReturnValue(throwError({ error: { params: { errMsg: 'fail' } } }))
      await component.uploadExcelFile(file)
      expect(loaderService.setLoaderState).toHaveBeenCalledWith(false)
      expect(snackbar.open).toHaveBeenCalledWith('fail')
    })

    it('should not show message on upload failure without errMsg', async () => {
      orgHieService.uploadFreameworkTemplate.mockReturnValue(throwError({}))
      await component.uploadExcelFile(file)
      expect(snackbar.open).not.toHaveBeenCalled()
    })

    it('should do nothing when response has no file name', async () => {
      orgHieService.uploadFreameworkTemplate.mockReturnValue(of({ result: {} }))
      const spy = jest.spyOn(component, 'getBulkuploadPrgressData').mockImplementation(() => { })
      await component.uploadExcelFile(file)
      expect(spy).not.toHaveBeenCalled()
    })
  })

  it('isValidExcelFile should validate the mime type', () => {
    expect(component.isValidExcelFile({ type: XLSX_TYPE } as any)).toBe(true)
    expect(component.isValidExcelFile({ type: 'text/csv' } as any)).toBe(false)
  })

  describe('handleDownloadFile', () => {
    let createObjectURL: any
    let revokeObjectURL: any
    let clickSpy: any

    beforeEach(() => {
      createObjectURL = jest.fn().mockReturnValue('blob:url')
      revokeObjectURL = jest.fn()
      ;(window.URL as any).createObjectURL = createObjectURL
      ;(window.URL as any).revokeObjectURL = revokeObjectURL
      clickSpy = jest.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => { })
    })

    it('should download file log', () => {
      component.handleDownloadFile({ fileName: 'a.xlsx' })
      expect(orgHieService.downloadFileLog).toHaveBeenCalledWith('a.xlsx')
      expect(createObjectURL).toHaveBeenCalled()
      expect(clickSpy).toHaveBeenCalled()
      expect(revokeObjectURL).toHaveBeenCalledWith('blob:url')
      expect(loaderService.setLoaderState).toHaveBeenCalledWith(false)
    })

    it('should show error when download fails', () => {
      orgHieService.downloadFileLog.mockReturnValue(throwError({}))
      component.handleDownloadFile({ fileName: 'a.xlsx' })
      expect(loaderService.setLoaderState).toHaveBeenCalledWith(false)
      expect(snackbar.open).toHaveBeenCalledWith('Error downloading file')
    })

    it('should show message when no file name provided', () => {
      component.handleDownloadFile({})
      expect(snackbar.open).toHaveBeenCalledWith('No file name provided for download')
    })

    it('should show message when item is null', async () => {
      component.handleDownloadFile(null)
      await flush()
      expect(snackbar.open).toHaveBeenCalledWith('No file name provided for download')
    })
  })
})
