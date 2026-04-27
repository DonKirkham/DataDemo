// ABOUTME: Idempotent provisioning of the Conference Events list, content types, and fields.
// ABOUTME: Uses PnPjs v4 for everything supported; raw REST POST for content-type field-link binding.

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/content-types';
import '@pnp/sp/content-types/list';
import '@pnp/sp/fields';
import { Logger, LogLevel } from '@pnp/logging';

import { IProvisioningService, IProvisioningSummary } from './IProvisioningService';
import { ConferenceEventsContentTypeIds } from '../models/conferenceEvents/ContentTypeIds';
import { SESSION_TYPES } from '../models/conferenceEvents/SessionType';

const LIST_TITLE = 'Conference Events';
const LIST_DESCRIPTION = 'Conferences, calls for speakers/sponsors, and sessions.';
const SITE_COLUMN_GROUP = 'Conference Events';
const CT_GROUP = 'Conference Events';

const SUBMISSION_STATUSES = ['Draft', 'Submitted', 'Accepted', 'Declined'];

interface IFieldRef {
  internalName: string;
  required?: boolean;
}

interface IContentTypeSpec {
  id: string;
  name: string;
  description: string;
  fields: IFieldRef[];
}

const CONTENT_TYPE_SPECS: IContentTypeSpec[] = [
  {
    id: ConferenceEventsContentTypeIds.Conference,
    name: 'Conference',
    description: 'A conference event with start/end dates and a card image.',
    fields: [
      { internalName: 'StartDate', required: true },
      { internalName: 'EndDate', required: true },
      { internalName: 'SiteUrl' },
      { internalName: 'CardImage' }
    ]
  },
  {
    id: ConferenceEventsContentTypeIds.CallForSpeakers,
    name: 'Call for Speakers',
    description: 'A call for speakers tied to a conference.',
    fields: [
      { internalName: 'StartDate', required: true },
      { internalName: 'EndDate', required: true },
      { internalName: 'SubmittedOn' },
      { internalName: 'SubmissionStatus', required: true }
    ]
  },
  {
    id: ConferenceEventsContentTypeIds.CallForSponsors,
    name: 'Call for Sponsors',
    description: 'A call for sponsors tied to a conference.',
    fields: [
      { internalName: 'StartDate', required: true },
      { internalName: 'EndDate', required: true },
      { internalName: 'SubmittedOn' },
      { internalName: 'SubmissionStatus', required: true }
    ]
  },
  {
    id: ConferenceEventsContentTypeIds.Session,
    name: 'Session',
    description: 'A conference session delivered by a speaker.',
    fields: [
      { internalName: 'SessionDateTime', required: true },
      { internalName: 'SessionType', required: true },
      { internalName: 'Speaker' }
    ]
  }
];

export class ProvisioningService implements IProvisioningService {
  constructor(private sp: SPFI, private siteUrl: string) {}

  public async isProvisioned(): Promise<boolean> {
    return this.listExists(LIST_TITLE);
  }

  public async ensureSchema(): Promise<IProvisioningSummary> {
    const summary: IProvisioningSummary = { created: [], existed: [] };

    Logger.write(`[Provisioning] Ensuring schema on ${this.siteUrl}`, LogLevel.Info);

    await this.ensureSiteColumns(summary);
    await this.ensureContentTypes(summary);
    await this.ensureFieldLinksOnContentTypes(summary);
    const list = await this.ensureList(summary);
    await this.ensureLookupFieldOnList(list, summary);
    await this.ensureContentTypesOnList(summary);
    await this.removeDefaultItemContentType(summary);

    Logger.write(
      `[Provisioning] Done. Created: ${summary.created.length}, Existed: ${summary.existed.length}`,
      LogLevel.Info
    );
    return summary;
  }

  private async ensureSiteColumns(summary: IProvisioningSummary): Promise<void> {
    await this.ensureField(summary, 'StartDate', () =>
      this.sp.web.fields.addDateTime('StartDate', { Group: SITE_COLUMN_GROUP }));

    await this.ensureField(summary, 'EndDate', () =>
      this.sp.web.fields.addDateTime('EndDate', { Group: SITE_COLUMN_GROUP }));

    await this.ensureField(summary, 'SiteUrl', () =>
      this.sp.web.fields.addUrl('SiteUrl', { Group: SITE_COLUMN_GROUP }));

    await this.ensureField(summary, 'CardImage', () =>
      this.sp.web.fields.addUrl('CardImage', { Group: SITE_COLUMN_GROUP }));

    await this.ensureField(summary, 'SubmittedOn', () =>
      this.sp.web.fields.addDateTime('SubmittedOn', { Group: SITE_COLUMN_GROUP }));

    await this.ensureField(summary, 'SubmissionStatus', () =>
      this.sp.web.fields.addChoice('SubmissionStatus', {
        Group: SITE_COLUMN_GROUP,
        Choices: SUBMISSION_STATUSES
      }));

    await this.ensureField(summary, 'SessionDateTime', () =>
      this.sp.web.fields.addDateTime('SessionDateTime', { Group: SITE_COLUMN_GROUP }));

    await this.ensureField(summary, 'SessionType', () =>
      this.sp.web.fields.addChoice('SessionType', {
        Group: SITE_COLUMN_GROUP,
        Choices: [...SESSION_TYPES]
      }));

    await this.ensureField(summary, 'Speaker', () =>
      this.sp.web.fields.addUser('Speaker', { Group: SITE_COLUMN_GROUP }));
  }

  private async ensureField(
    summary: IProvisioningSummary,
    internalName: string,
    create: () => Promise<unknown>
  ): Promise<void> {
    if (await this.fieldExists(internalName)) {
      summary.existed.push(`field:${internalName}`);
      return;
    }
    await create();
    summary.created.push(`field:${internalName}`);
    Logger.write(`[Provisioning] Created site column ${internalName}`, LogLevel.Info);
  }

  private async fieldExists(internalName: string): Promise<boolean> {
    try {
      await this.sp.web.fields.getByInternalNameOrTitle(internalName)();
      return true;
    } catch {
      return false;
    }
  }

  private async ensureContentTypes(summary: IProvisioningSummary): Promise<void> {
    for (const spec of CONTENT_TYPE_SPECS) {
      if (await this.contentTypeExists(spec.id)) {
        summary.existed.push(`ct:${spec.name}`);
        continue;
      }
      await this.sp.web.contentTypes.add(spec.id, spec.name, spec.description, CT_GROUP);
      summary.created.push(`ct:${spec.name}`);
      Logger.write(`[Provisioning] Created content type ${spec.name}`, LogLevel.Info);
    }
  }

  private async contentTypeExists(id: string): Promise<boolean> {
    try {
      await this.sp.web.contentTypes.getById(id)();
      return true;
    } catch {
      return false;
    }
  }

  private async ensureFieldLinksOnContentTypes(summary: IProvisioningSummary): Promise<void> {
    for (const spec of CONTENT_TYPE_SPECS) {
      const existingLinks = await this.sp.web.contentTypes.getById(spec.id).fieldLinks();
      const existingNames = new Set(existingLinks.map(fl => fl.Name));

      for (const ref of spec.fields) {
        if (existingNames.has(ref.internalName)) {
          summary.existed.push(`fieldLink:${spec.name}:${ref.internalName}`);
          continue;
        }
        await this.addFieldLink(spec.id, ref.internalName, ref.required === true);
        summary.created.push(`fieldLink:${spec.name}:${ref.internalName}`);
        Logger.write(
          `[Provisioning] Linked ${ref.internalName} -> ${spec.name}`,
          LogLevel.Info
        );
      }
    }
  }

  private async addFieldLink(
    contentTypeId: string,
    fieldInternalName: string,
    required: boolean
  ): Promise<void> {
    const url = `${this.trimTrailingSlash(this.siteUrl)}/_api/web/contenttypes('${encodeURIComponent(contentTypeId)}')/fieldlinks`;
    const digest = await this.getRequestDigest();

    const body = {
      __metadata: { type: 'SP.FieldLink' },
      FieldInternalName: fieldInternalName,
      Required: required
    };

    const resp = await fetch(url, {
      method: 'POST',
      credentials: 'include',
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': digest,
        'odata-version': ''
      },
      body: JSON.stringify(body)
    });

    if (!resp.ok) {
      const text = await resp.text();
      throw new Error(`Failed to add field link ${fieldInternalName} to ${contentTypeId}: ${resp.status} ${text}`);
    }
  }

  private async getRequestDigest(): Promise<string> {
    const url = `${this.trimTrailingSlash(this.siteUrl)}/_api/contextinfo`;
    const resp = await fetch(url, {
      method: 'POST',
      credentials: 'include',
      headers: {
        Accept: 'application/json;odata=verbose'
      }
    });
    if (!resp.ok) {
      throw new Error(`Failed to fetch request digest: ${resp.status}`);
    }
    const json = await resp.json();
    return json.d.GetContextWebInformation.FormDigestValue;
  }

  private trimTrailingSlash(url: string): string {
    return url.endsWith('/') ? url.slice(0, -1) : url;
  }

  private async ensureList(summary: IProvisioningSummary): Promise<{ Id: string }> {
    if (await this.listExists(LIST_TITLE)) {
      summary.existed.push(`list:${LIST_TITLE}`);
      const info = await this.sp.web.lists.getByTitle(LIST_TITLE).select('Id')();
      // PnPjs returns Id as string GUID
      return { Id: (info as { Id: string }).Id };
    }

    const ensureResult = await this.sp.web.lists.ensure(
      LIST_TITLE,
      LIST_DESCRIPTION,
      100,
      true
    );
    summary.created.push(`list:${LIST_TITLE}`);
    Logger.write(`[Provisioning] Created list ${LIST_TITLE}`, LogLevel.Info);

    const info = await ensureResult.list.select('Id')();
    return { Id: (info as { Id: string }).Id };
  }

  private async listExists(title: string): Promise<boolean> {
    try {
      await this.sp.web.lists.getByTitle(title).select('Id')();
      return true;
    } catch {
      return false;
    }
  }

  private async ensureLookupFieldOnList(
    list: { Id: string },
    summary: IProvisioningSummary
  ): Promise<void> {
    const fieldName = 'ConferenceLookup';
    const listFields = this.sp.web.lists.getByTitle(LIST_TITLE).fields;

    try {
      await listFields.getByInternalNameOrTitle(fieldName)();
      summary.existed.push(`field:${fieldName}`);
      return;
    } catch {
      // not present yet
    }

    await listFields.addLookup(fieldName, {
      LookupListId: list.Id,
      LookupFieldName: 'Title'
    });
    summary.created.push(`field:${fieldName}`);
    Logger.write(`[Provisioning] Created lookup field ${fieldName}`, LogLevel.Info);

    // Bind ConferenceLookup as a field link on the three child content types.
    for (const spec of CONTENT_TYPE_SPECS) {
      if (spec.name === 'Conference') continue;
      const existingLinks = await this.sp.web.contentTypes.getById(spec.id).fieldLinks();
      if (existingLinks.some(fl => fl.Name === fieldName)) {
        summary.existed.push(`fieldLink:${spec.name}:${fieldName}`);
        continue;
      }
      await this.addFieldLink(spec.id, fieldName, false);
      summary.created.push(`fieldLink:${spec.name}:${fieldName}`);
    }
  }

  private async ensureContentTypesOnList(summary: IProvisioningSummary): Promise<void> {
    const list = this.sp.web.lists.getByTitle(LIST_TITLE);
    await list.update({ ContentTypesEnabled: true });

    const existing = await list.contentTypes();
    const existingIds = new Set(existing.map(ct => ct.StringId));

    for (const spec of CONTENT_TYPE_SPECS) {
      const alreadyOnList = Array.from(existingIds).some(id => id.startsWith(spec.id));
      if (alreadyOnList) {
        summary.existed.push(`listCt:${spec.name}`);
        continue;
      }
      await list.contentTypes.addAvailableContentType(spec.id);
      summary.created.push(`listCt:${spec.name}`);
      Logger.write(`[Provisioning] Bound ${spec.name} to list`, LogLevel.Info);
    }
  }

  private async removeDefaultItemContentType(summary: IProvisioningSummary): Promise<void> {
    const list = this.sp.web.lists.getByTitle(LIST_TITLE);
    const cts = await list.contentTypes();
    const itemCt = cts.find(ct => ct.StringId.startsWith('0x0100') === false && ct.StringId.startsWith('0x01'));
    if (!itemCt) {
      summary.existed.push('listCt:Item(absent)');
      return;
    }
    await list.contentTypes.getById(itemCt.StringId).delete();
    summary.created.push('listCt:Item(removed)');
    Logger.write('[Provisioning] Removed default Item content type from list', LogLevel.Info);
  }
}
