import { of, Subject, throwError } from 'rxjs'
import { OrgHierarchyMappingComponent } from './org-hierarchy-mapping.component'
import { BulkUploadOrgComponent } from '../../bulk-upload-org/bulk-upload-org.component'

const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

describe('OrgHierarchyMappingComponent', () => {
  let component: OrgHierarchyMappingComponent
  let snackbar: any
  let orgHieService: any
  let loaderService: any
  let router: any
  let activatedRoute: any
  let dialog: any

  const orgs: any[] = [
    { identifier: 'o1', orgName: 'Alpha', orgHierarchyFrameworkId: 'fw1' },
    { identifier: 'o2', orgName: 'Beta' },
  ]

  const flush = () => new Promise(resolve => setTimeout(resolve, 0))

  function createComponent(userRoles?: Set<string>) {
    activatedRoute = {
      snapshot: { parent: { data: { configService: { userRoles, userProfile: { rootOrgId: 'root-1' } } } } },
    }
    component = new OrgHierarchyMappingComponent(
      snackbar as any,
      orgHieService as any,
      loaderService as any,
      router as any,
      activatedRoute as any,
      dialog as any
    )
  }

  beforeEach(() => {
    snackbar = { open: jest.fn() }
    orgHieService = {
      getCenterOrStateList: jest.fn().mockReturnValue(of({ result: { response: { content: orgs } } })),
      getOrgReadData: jest.fn().mockReturnValue(of({
        params: { status: 'SUCCESS' },
        result: {
          response: { id: 'state-1', orgName: 'StateOrg', sbOrgType: 'state', ministryOrStateId: 'ms-1', identifier: 'state-1' },
        },
      })),
      getOrganizationDetails: jest.fn().mockReturnValue(of({ result: { content: [{ id: 'child' }] } })),
      createMasterFrameWork: jest.fn().mockReturnValue(of({ result: { message: 'created' } })),
      downloadSampleTemplate: jest.fn().mockReturnValue(of({})),
      exportFramework: jest.fn().mockReturnValue(of({})),
      uploadFreameworkTemplate: jest.fn().mockReturnValue(of({ result: { fileName: 'f.xlsx' } })),
    }
    loaderService = { setLoaderState: jest.fn() }
    router = { navigate: jest.fn() }
    dialog = { open: jest.fn().mockReturnValue({ afterClosed: () => of(true) }) }
    jest.spyOn(console, 'log').mockImplementation(() => { })
    jest.spyOn(console, 'error').mockImplementation(() => { })
    createComponent(new Set(['spv_admin']))
  })

  afterEach(() => {
    jest.useRealTimers()
    jest.restoreAllMocks()
  })

  it('should create', () => {
    expect(component).toBeTruthy()
  })

  describe('getters', () => {
    it('should return user roles and org id from route', () => {
      expect(component.userRoles).toEqual(new Set(['spv_admin']))
      expect(component.orgId).toBe('root-1')
    })
  })

  describe('checkIfStateAdmin', () => {
    it('should return true only for state admin without spv admin', () => {
      expect(component.checkIfStateAdmin()).toBe(false)
      createComponent(new Set(['state_admin']))
      expect(component.checkIfStateAdmin()).toBe(true)
      createComponent(new Set(['state_admin', 'spv_admin']))
      expect(component.checkIfStateAdmin()).toBe(false)
      createComponent(undefined)
      expect(component.checkIfStateAdmin()).toBeFalsy()
    })
  })

  describe('ngOnInit', () => {
    it('should fetch org list for default type as SPV admin', async () => {
      component.ngOnInit()
      await flush()
      expect(orgHieService.getCenterOrStateList).toHaveBeenCalledWith(
        expect.objectContaining({ request: expect.objectContaining({ filters: { status: 1, sbOrgType: 'state' }, query: '' }) })
      )
      expect(component.filteredOrganizations).toEqual(orgs)
    })

    it('should search after debounce when search value changes', () => {
      jest.useFakeTimers()
      const spy = jest.spyOn(component, 'getCentenrOrStateList').mockResolvedValue(undefined as any)
      component.ngOnInit()
      spy.mockClear()
      component.searchControl.setValue('alp')
      jest.advanceTimersByTime(700)
      expect(spy).toHaveBeenCalledWith('state', 'alp')
    })

    it('should not search when no org type selected', () => {
      jest.useFakeTimers()
      component.selectedOrgType = ''
      const spy = jest.spyOn(component, 'getCentenrOrStateList').mockResolvedValue(undefined as any)
      component.ngOnInit()
      component.searchControl.setValue('alp')
      jest.advanceTimersByTime(700)
      expect(spy).not.toHaveBeenCalled()
    })

    it('should call getOrgReadAndDetails for state admin', () => {
      createComponent(new Set(['state_admin']))
      const spy = jest.spyOn(component, 'getOrgReadAndDetails').mockImplementation(() => { })
      component.ngOnInit()
      expect(spy).toHaveBeenCalled()
      expect(orgHieService.getCenterOrStateList).not.toHaveBeenCalled()
    })

    it('should recreate tree view when organization changes', () => {
      jest.useFakeTimers()
      jest.spyOn(component, 'getCentenrOrStateList').mockResolvedValue(undefined as any)
      component.ngOnInit()
      component.organizationCtrl.setValue('o1')
      expect(component.showTreeView).toBe(false)
      jest.advanceTimersByTime(0)
      expect(component.showTreeView).toBe(true)
    })
  })

  it('ngOnDestroy should complete destroy subject', () => {
    const destroy$: any = (component as any).destroy$
    const nextSpy = jest.spyOn(destroy$, 'next')
    const completeSpy = jest.spyOn(destroy$, 'complete')
    component.ngOnDestroy()
    expect(nextSpy).toHaveBeenCalled()
    expect(completeSpy).toHaveBeenCalled()
  })

  describe('ngAfterViewInit', () => {
    it('should focus search input when opened and reset search when closed', () => {
      jest.useFakeTimers()
      const openedChange = new Subject<boolean>()
      const focus = jest.fn()
      component.singleSelect = { openedChange } as any
      component.searchInput = { nativeElement: { focus } } as any
      component.ngAfterViewInit()

      openedChange.next(true)
      jest.advanceTimersByTime(0)
      expect(focus).toHaveBeenCalled()

      component.searchControl.setValue('abc')
      openedChange.next(false)
      expect(component.searchControl.value).toBe('')
    })

    it('should not fail when singleSelect is undefined', () => {
      component.singleSelect = undefined as any
      expect(() => component.ngAfterViewInit()).not.toThrow()
    })
  })

  describe('filterOrganizations', () => {
    it('should filter by name and reset on empty value', () => {
      component.allOrganizations = orgs as any
      component.filterOrganizations('alp')
      expect(component.filteredOrganizations).toEqual([orgs[0]])
      component.filterOrganizations('')
      expect(component.filteredOrganizations).toEqual(orgs)
    })
  })

  describe('orgSelected', () => {
    it('should refetch when org type changes', () => {
      const spy = jest.spyOn(component, 'getCentenrOrStateList').mockResolvedValue(undefined as any)
      component.organizationCtrl.setValue('o1')
      component.orgSelected('ministry')
      expect(component.organizationCtrl.value).toBeNull()
      expect(component.selectedOrgType).toBe('ministry')
      expect(spy).toHaveBeenCalledWith('ministry')
    })

    it('should not refetch when same org type selected', () => {
      const spy = jest.spyOn(component, 'getCentenrOrStateList').mockResolvedValue(undefined as any)
      component.orgSelected('state')
      expect(spy).not.toHaveBeenCalled()
    })
  })

  describe('getCentenrOrStateList', () => {
    it('should request ministry orgs', async () => {
      await component.getCentenrOrStateList('ministry', 'abc')
      expect(orgHieService.getCenterOrStateList).toHaveBeenCalledWith(
        expect.objectContaining({ request: expect.objectContaining({ filters: { status: 1, sbOrgType: 'ministry' }, query: 'abc' }) })
      )
      expect(component.allOrganizations).toEqual(orgs)
    })

    it('should keep empty org type for unknown type', async () => {
      await component.getCentenrOrStateList('other')
      expect(orgHieService.getCenterOrStateList).toHaveBeenCalledWith(
        expect.objectContaining({ request: expect.objectContaining({ filters: { status: 1, sbOrgType: '' } }) })
      )
    })

    it('should clear lists when no organisations returned', async () => {
      orgHieService.getCenterOrStateList.mockReturnValue(of({ result: { response: { content: [] } } }))
      await component.getCentenrOrStateList('state')
      expect(component.allOrganizations).toEqual([])
      expect(component.filteredOrganizations).toEqual([])
    })

    it('should clear lists on API error', async () => {
      orgHieService.getCenterOrStateList.mockReturnValue(throwError({}))
      await component.getCentenrOrStateList('state')
      expect(loaderService.setLoaderState).toHaveBeenCalledWith(false)
      expect(component.allOrganizations).toEqual([])
    })
  })

  describe('getOrgDetails', () => {
    it('should return null when no organization selected', () => {
      expect(component.getOrgDetails()).toBeNull()
    })

    it('should return null when filtered list empty for non state admin', () => {
      component.organizationCtrl.setValue('o1')
      component.filteredOrganizations = []
      expect(component.getOrgDetails()).toBeNull()
    })

    it('should return selected org for SPV admin', () => {
      component.filteredOrganizations = orgs
      component.organizationCtrl.setValue('o1')
      expect(component.getOrgDetails()).toEqual(orgs[0])
      component.organizationCtrl.setValue('x')
      expect(component.getOrgDetails()).toBeNull()
    })

    it('should return org read data for state admin', () => {
      createComponent(new Set(['state_admin']))
      component.orgReadData = { id: 's1' }
      component.organizationCtrl.setValue('s1')
      expect(component.getOrgDetails()).toEqual({ id: 's1' })
    })
  })

  describe('hasOrgHierarchyFrameworkId', () => {
    it('should return false when no org selected', () => {
      expect(component.hasOrgHierarchyFrameworkId()).toBe(false)
    })

    it('should check framework id for SPV admin', () => {
      component.filteredOrganizations = orgs
      component.organizationCtrl.setValue('o1')
      expect(component.hasOrgHierarchyFrameworkId()).toBe(true)
      component.organizationCtrl.setValue('o2')
      expect(component.hasOrgHierarchyFrameworkId()).toBe(false)
    })

    it('should check framework id for state admin', () => {
      createComponent(new Set(['state_admin']))
      component.orgReadData = { orgHierarchyFrameworkId: 'fw' }
      component.organizationCtrl.setValue('s1')
      expect(component.hasOrgHierarchyFrameworkId()).toBe(true)
    })
  })

  it('cancelHierarchyCreation should reset selection and close select', () => {
    const close = jest.fn()
    component.singleSelect = { close } as any
    component.allOrganizations = orgs as any
    component.organizationCtrl.setValue('o1')
    component.cancelHierarchyCreation()
    expect(component.organizationCtrl.value).toBeNull()
    expect(component.filteredOrganizations).toEqual(orgs)
    expect(close).toHaveBeenCalled()
  })

  describe('createNewHierarchy', () => {
    beforeEach(() => {
      component.singleSelect = { close: jest.fn() } as any
    })

    it('should create framework for SPV admin', async () => {
      jest.useFakeTimers()
      const listSpy = jest.spyOn(component, 'getCentenrOrStateList').mockResolvedValue(undefined as any)
      component.allOrganizations = orgs as any
      component.filteredOrganizations = orgs
      component.organizationCtrl.setValue('o2')
      await component.createNewHierarchy()
      expect(orgHieService.createMasterFrameWork).toHaveBeenCalledWith({ request: { parentOrgId: 'o2' } })
      jest.advanceTimersByTime(2000)
      expect(listSpy).toHaveBeenCalledWith('state')
      expect(component.organizationCtrl.value).toBe('o2')
      expect(snackbar.open).toHaveBeenCalledWith('Framework created successfully for Beta')
    })

    it('should create framework for state admin', async () => {
      jest.useFakeTimers()
      createComponent(new Set(['state_admin']))
      component.singleSelect = { close: jest.fn() } as any
      const listSpy = jest.spyOn(component, 'getCentenrOrStateList')
      component.orgReadData = { id: 'state-1', orgName: 'StateOrg' }
      await component.createNewHierarchy()
      expect(orgHieService.createMasterFrameWork).toHaveBeenCalledWith({ request: { parentOrgId: 'state-1' } })
      jest.advanceTimersByTime(2000)
      expect(listSpy).not.toHaveBeenCalled()
      expect(snackbar.open).toHaveBeenCalledWith('Framework created successfully for StateOrg')
    })

    it('should show error message when creation throws with errMsg', async () => {
      orgHieService.createMasterFrameWork.mockReturnValue(throwError({ params: { errMsg: 'boom' } }))
      component.filteredOrganizations = orgs
      component.organizationCtrl.setValue('o2')
      await component.createNewHierarchy()
      expect(snackbar.open).toHaveBeenCalledWith('boom')
      expect(snackbar.open).toHaveBeenCalledWith('Failed to create framework for Beta')
    })

    it('should show failure message when creation throws without errMsg', async () => {
      orgHieService.createMasterFrameWork.mockReturnValue(throwError({}))
      component.filteredOrganizations = orgs
      component.organizationCtrl.setValue('o2')
      await component.createNewHierarchy()
      expect(snackbar.open).toHaveBeenCalledTimes(1)
      expect(snackbar.open).toHaveBeenCalledWith('Failed to create framework for Beta')
    })

    it('should show failure message when response has no message', async () => {
      orgHieService.createMasterFrameWork.mockReturnValue(of({ result: {} }))
      component.filteredOrganizations = orgs
      component.organizationCtrl.setValue('o2')
      await component.createNewHierarchy()
      expect(snackbar.open).toHaveBeenCalledWith('Failed to create framework for Beta')
    })

    it('should do nothing when no org selected', async () => {
      await component.createNewHierarchy()
      expect(orgHieService.createMasterFrameWork).not.toHaveBeenCalled()
    })
  })

  it('checkloader should set loader state', () => {
    component.checkloader(true)
    expect(loaderService.setLoaderState).toHaveBeenCalledWith(true)
  })

  describe('redirectOrg', () => {
    it('should navigate with org details for SPV admin', async () => {
      orgHieService.getCenterOrStateList.mockReturnValue(of({ result: { response: { content: [{ organisationType: 2 }] } } }))
      await component.redirectOrg({ name: 'Alpha', additionalProperties: { orgId: 'o1' } })
      expect(router.navigate).toHaveBeenCalledWith(['/app/roles/o1/users'], {
        queryParams: expect.objectContaining({ roleId: 'o1', subOrgType: 'state', organisationType: 2, orgName: 'Alpha' }),
      })
    })

    it('should navigate with state sub org type for state admin and no details', async () => {
      createComponent(new Set(['state_admin']))
      orgHieService.getCenterOrStateList.mockReturnValue(of({}))
      await component.redirectOrg({ name: 'S', additionalProperties: { orgId: 's1' } })
      expect(router.navigate).toHaveBeenCalledWith(['/app/roles/s1/users'], {
        queryParams: expect.objectContaining({ subOrgType: 'state', organisationType: undefined }),
      })
    })
  })

  describe('readOrganizationDetails', () => {
    it('should return content when available', async () => {
      const res = await component.readOrganizationDetails('o1')
      expect(res).toEqual(orgs)
    })

    it('should return null when no content', async () => {
      orgHieService.getCenterOrStateList.mockReturnValue(of(null))
      const res = await component.readOrganizationDetails('o1')
      expect(res).toBeNull()
    })
  })

  describe('downloadTemplate', () => {
    beforeEach(() => {
      component.allOrganizations = orgs as any
      component.organizationCtrl.setValue('o1')
    })

    it('should download template successfully', async () => {
      await component.downloadTemplate()
      expect(orgHieService.downloadSampleTemplate).toHaveBeenCalledWith('o1')
      expect(snackbar.open).toHaveBeenCalledWith('Download successfully')
    })

    it('should show error message on failure', async () => {
      orgHieService.downloadSampleTemplate.mockReturnValue(throwError({ error: { params: { errMsg: 'dl err' } } }))
      await component.downloadTemplate()
      expect(loaderService.setLoaderState).toHaveBeenCalledWith(false)
      expect(snackbar.open).toHaveBeenCalledWith('dl err')
    })

    it('should not show message on failure without errMsg', async () => {
      orgHieService.downloadSampleTemplate.mockReturnValue(throwError({}))
      await component.downloadTemplate()
      expect(snackbar.open).not.toHaveBeenCalled()
    })

    it('should do nothing when no org selected', async () => {
      component.organizationCtrl.setValue('none')
      await component.downloadTemplate()
      expect(orgHieService.downloadSampleTemplate).not.toHaveBeenCalled()
    })
  })

  describe('exportData', () => {
    beforeEach(() => {
      component.allOrganizations = orgs as any
      component.organizationCtrl.setValue('o1')
    })

    it('should export successfully', async () => {
      await component.exportData()
      expect(orgHieService.exportFramework).toHaveBeenCalledWith('o1')
      expect(snackbar.open).toHaveBeenCalledWith('Exported successfully for Alpha')
    })

    it('should show error message on failure', async () => {
      orgHieService.exportFramework.mockReturnValue(throwError({ error: { params: { errMsg: 'ex err' } } }))
      await component.exportData()
      expect(loaderService.setLoaderState).toHaveBeenCalledWith(false)
      expect(snackbar.open).toHaveBeenCalledWith('ex err')
    })

    it('should not show message on failure without errMsg', async () => {
      orgHieService.exportFramework.mockReturnValue(throwError({}))
      await component.exportData()
      expect(snackbar.open).not.toHaveBeenCalled()
    })

    it('should do nothing when no org selected', async () => {
      component.organizationCtrl.setValue('none')
      await component.exportData()
      expect(orgHieService.exportFramework).not.toHaveBeenCalled()
    })
  })

  describe('getselectedOrgData', () => {
    it('should return matching org for SPV admin or null', () => {
      component.allOrganizations = orgs as any
      component.organizationCtrl.setValue('o2')
      expect(component.getselectedOrgData()).toEqual(orgs[1])
      component.organizationCtrl.setValue('none')
      expect(component.getselectedOrgData()).toBeNull()
    })

    it('should return org read data for state admin', () => {
      createComponent(new Set(['state_admin']))
      component.orgReadData = { id: 's1' }
      expect(component.getselectedOrgData()).toEqual({ id: 's1' })
    })
  })

  describe('onFileSelected', () => {
    let nativeElement: any

    beforeEach(() => {
      nativeElement = { value: 'file.txt' }
      component.fileInput = { nativeElement } as any
    })

    it('should reject invalid file types and clear input', () => {
      const spy = jest.spyOn(component, 'uploadExcelFile')
      component.onFileSelected({ target: { files: [{ type: 'text/plain', size: 1 }] } })
      expect(spy).not.toHaveBeenCalled()
      expect(snackbar.open).toHaveBeenCalledWith('Please select a valid Excel file (.xlsx)', 'Close', { duration: 5000 })
      expect(nativeElement.value).toBe('')
    })

    it('should reject files larger than 5MB', () => {
      const spy = jest.spyOn(component, 'uploadExcelFile')
      component.onFileSelected({ target: { files: [{ type: XLSX_TYPE, size: 6 * 1024 * 1024 }] } })
      expect(spy).not.toHaveBeenCalled()
      expect(snackbar.open).toHaveBeenCalledWith('File size should not exceed 5MB', 'Close', { duration: 5000 })
    })

    it('should upload valid file', () => {
      const spy = jest.spyOn(component, 'uploadExcelFile').mockResolvedValue(undefined as any)
      const file: any = { type: XLSX_TYPE, size: 1 }
      component.onFileSelected({ target: { files: [file] } })
      expect(spy).toHaveBeenCalledWith(file)
    })

    it('should do nothing when no file selected', () => {
      const spy = jest.spyOn(component, 'uploadExcelFile')
      component.onFileSelected({ target: { files: [] } })
      expect(spy).not.toHaveBeenCalled()
    })
  })

  it('clearFileInput should do nothing without file input', () => {
    component.fileInput = undefined as any
    expect(() => component.clearFileInput()).not.toThrow()
  })

  describe('uploadExcelFile', () => {
    const file = new File(['x'], 'f.xlsx', { type: XLSX_TYPE })

    beforeEach(() => {
      component.allOrganizations = orgs as any
      component.organizationCtrl.setValue('o1')
    })

    it('should upload successfully', async () => {
      await component.uploadExcelFile(file)
      expect(orgHieService.uploadFreameworkTemplate).toHaveBeenCalledWith(expect.any(FormData), orgs[0])
      expect(component.bulkUploadRefresh).toBe(true)
      expect(snackbar.open).toHaveBeenCalledWith('File uploaded successfully. Please check after 5 minutes for the results.')
    })

    it('should show error on upload failure', async () => {
      orgHieService.uploadFreameworkTemplate.mockReturnValue(throwError({ error: { params: { errMsg: 'up err' } } }))
      await component.uploadExcelFile(file)
      expect(component.bulkUploadRefresh).toBe(false)
      expect(snackbar.open).toHaveBeenCalledWith('up err')
    })

    it('should not show message on failure without errMsg', async () => {
      orgHieService.uploadFreameworkTemplate.mockReturnValue(throwError({}))
      await component.uploadExcelFile(file)
      expect(snackbar.open).not.toHaveBeenCalled()
    })
  })

  describe('getOrgReadAndDetails', () => {
    beforeEach(() => {
      createComponent(new Set(['state_admin']))
    })

    it('should read org and fetch org details', () => {
      component.getOrgReadAndDetails()
      expect(orgHieService.getOrgReadData).toHaveBeenCalledWith({ request: { organisationId: 'root-1' } })
      expect(orgHieService.getOrganizationDetails).toHaveBeenCalledWith({
        request: { filters: { status: 1, ministryOrStateType: 'state', ministryOrStateId: 'ms-1' } },
      })
      expect(component.orgSearchData).toEqual([{ id: 'child' }])
      expect(component.organizationCtrl.value).toBe('ms-1')
      expect(loaderService.setLoaderState).toHaveBeenLastCalledWith(false)
    })

    it('should drop empty filters and use rootOrgId', () => {
      orgHieService.getOrgReadData.mockReturnValue(of({
        params: { status: 'success' },
        result: { response: { rootOrgId: 'r1' } },
      }))
      orgHieService.getOrganizationDetails.mockReturnValue(of({ result: {} }))
      component.getOrgReadAndDetails()
      expect(orgHieService.getOrganizationDetails).toHaveBeenCalledWith({ request: { filters: { status: 1 } } })
      expect(component.organizationCtrl.value).toBe('r1')
      expect(component.orgSearchData).toBeUndefined()
    })

    it('should stop when read response is empty', () => {
      orgHieService.getOrgReadData.mockReturnValue(of({ params: { status: 'success' }, result: {} }))
      component.getOrgReadAndDetails()
      expect(component.orgReadData).toBeNull()
      expect(orgHieService.getOrganizationDetails).not.toHaveBeenCalled()
    })

    it('should show error message when read fails', () => {
      orgHieService.getOrgReadData.mockReturnValue(of({
        params: { status: 'failed' },
        error: { params: { errMsg: 'read err' } },
      }))
      component.getOrgReadAndDetails()
      expect(snackbar.open).toHaveBeenCalledWith('read err')
      expect(orgHieService.getOrganizationDetails).not.toHaveBeenCalled()
    })

    it('should not show message when read fails without errMsg', () => {
      orgHieService.getOrgReadData.mockReturnValue(of(null))
      component.getOrgReadAndDetails()
      expect(snackbar.open).not.toHaveBeenCalled()
    })

    it('should show error message when API chain errors', () => {
      orgHieService.getOrganizationDetails.mockReturnValue(throwError({ error: { params: { errMsg: 'chain err' } } }))
      component.getOrgReadAndDetails()
      expect(snackbar.open).toHaveBeenCalledWith('chain err')
    })

    it('should not show message when API chain errors without errMsg', () => {
      orgHieService.getOrgReadData.mockReturnValue(throwError({}))
      component.getOrgReadAndDetails()
      expect(snackbar.open).not.toHaveBeenCalled()
      expect(loaderService.setLoaderState).toHaveBeenLastCalledWith(false)
    })
  })

  describe('openBulkUploadDialog', () => {
    it('should open dialog and refresh list after close for SPV admin', async () => {
      const listSpy = jest.spyOn(component, 'getCentenrOrStateList').mockResolvedValue(undefined as any)
      component.allOrganizations = orgs as any
      component.organizationCtrl.setValue('o1')
      component.openBulkUploadDialog()
      expect(dialog.open).toHaveBeenCalledWith(BulkUploadOrgComponent, expect.objectContaining({
        data: { bulkUploadConfig: expect.objectContaining({ frameworkData: orgs[0] }) },
      }))
      await flush()
      expect(listSpy).toHaveBeenCalledWith('state')
      expect(component.bulkUploadRefresh).toBe(false)
    })

    it('should refresh org read data after close for state admin', async () => {
      createComponent(new Set(['state_admin']))
      const readSpy = jest.spyOn(component, 'getOrgReadAndDetails').mockImplementation(() => { })
      component.openBulkUploadDialog()
      await flush()
      expect(readSpy).toHaveBeenCalled()
      expect(component.bulkUploadRefresh).toBe(false)
    })
  })

  it('recreateTreeView should toggle tree view', () => {
    jest.useFakeTimers()
    component.showTreeView = true
    component.recreateTreeView()
    expect(component.showTreeView).toBe(false)
    jest.advanceTimersByTime(0)
    expect(component.showTreeView).toBe(true)
  })
})
