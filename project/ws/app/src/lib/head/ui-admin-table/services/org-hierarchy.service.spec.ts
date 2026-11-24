import { of } from 'rxjs'
import { OrgHierarchyService } from './org-hierarchy.service'

describe('OrgHierarchyService', () => {
  let service: OrgHierarchyService
  let http: any

  beforeEach(() => {
    http = {
      get: jest.fn().mockReturnValue(of({ get: true })),
      post: jest.fn().mockReturnValue(of({ post: true })),
      patch: jest.fn().mockReturnValue(of({ patch: true })),
    }
    service = new OrgHierarchyService(http as any)
  })

  afterEach(() => {
    jest.clearAllMocks()
  })

  it('should be created', () => {
    expect(service).toBeTruthy()
  })

  it('getCenterOrStateList should POST to org search', done => {
    const req: any = { request: { filters: {} } }
    service.getCenterOrStateList(req).subscribe((res: any) => {
      expect(res).toEqual({ post: true })
      done()
    })
    expect(http.post).toHaveBeenCalledWith('/apis/proxies/v8/org/v1/search', req)
  })

  it('createMasterFrameWork should POST to hierarchy link', () => {
    const req: any = { request: { parentOrgId: 'p1' } }
    service.createMasterFrameWork(req)
    expect(http.post).toHaveBeenCalledWith('/apis/proxies/v8/org/v2/hierarchy/link', req)
  })

  it('downloadFileLog should GET the file log', () => {
    service.downloadFileLog('log.xlsx')
    expect(http.get).toHaveBeenCalledWith('/apis/proxies/v8/organisation/v1/hierarchy/download/file/log.xlsx')
  })

  it('downloadSampleTemplate should GET the sample template for root org', () => {
    service.downloadSampleTemplate('root1')
    expect(http.get).toHaveBeenCalledWith('/apis/proxies/v8/org/v2/hierarchy/bulk/link/root1/sample')
  })

  it('exportFramework should GET the export for root org', () => {
    service.exportFramework('root1')
    expect(http.get).toHaveBeenCalledWith('/apis/proxies/v8/org/v2/hierarchy/bulk/link/root1/export')
  })

  it('uploadFreameworkTemplate should POST form data with framework id', () => {
    const formData: any = { file: 'x' }
    service.uploadFreameworkTemplate(formData, { orgHierarchyFrameworkId: 'fw1' })
    expect(http.post).toHaveBeenCalledWith('/apis/proxies/v8/org/v2/hierarchy/bulk/link/fw1', formData)
  })

  it('getOrgReadData should POST to org read', () => {
    const req: any = { request: { organisationId: 'o1' } }
    service.getOrgReadData(req)
    expect(http.post).toHaveBeenCalledWith('/apis/proxies/v8/org/v1/read', req)
  })

  it('getOrganizationDetails should POST to org search', () => {
    const req: any = { request: { filters: { status: 1 } } }
    service.getOrganizationDetails(req)
    expect(http.post).toHaveBeenCalledWith('/apis/proxies/v8/org/v1/search', req)
  })

  it('getBulkuploadProgress should GET progress details', () => {
    service.getBulkuploadProgress('org1')
    expect(http.get).toHaveBeenCalledWith('/apis/proxies/v8/organisation/v1/hierarchy/progress/details/bulkUpload/org1')
  })

  it('setOrgStatus should PATCH status update', () => {
    const req: any = { request: { status: 0 } }
    service.setOrgStatus(req)
    expect(http.patch).toHaveBeenCalledWith('apis/proxies/v8/org/v1/status/update', req)
  })

  it('should set and get org data', () => {
    expect(service.getOrgData()).toBeUndefined()
    service.setOrgData({ id: 'o1' })
    expect(service.getOrgData()).toEqual({ id: 'o1' })
  })

  it('should set and get parent org data', () => {
    expect(service.getParentOrgData()).toBeUndefined()
    service.setParentOrgData({ id: 'p1' })
    expect(service.getParentOrgData()).toEqual({ id: 'p1' })
  })
})
