import { NextResponse } from "next/server";

import { withAdminApi, ValidationError } from "@/lib/api/adminApi";
import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  MAX_IMAGE_BYTES,
  uploadAndSetListImage,
} from "@/lib/services/listThumbnailService";
import { mapCompanyFields } from "@/lib/services/companyService";
import { getListItemByKey } from "@/lib/services/sharePointListService";

export const dynamic = "force-dynamic";

export async function POST(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const { id } = await context.params;
  return withAdminApi(
    "POST /api/admin/companies/[id]/logo",
    async (_ctx, req) => {
      const contentType = req.headers.get("content-type") ?? "";
      if (!contentType.toLowerCase().includes("multipart/form-data")) {
        throw new ValidationError("Expected multipart form data.");
      }
      const form = await req.formData();
      const file = form.get("file");
      if (!(file instanceof File)) {
        throw new ValidationError("Image file is required.");
      }
      if (file.size > MAX_IMAGE_BYTES) {
        throw new ValidationError("Image must be 10 MB or smaller.");
      }
      const bytes = new Uint8Array(await file.arrayBuffer());
      const fields = getSharePointFields("company");
      await uploadAndSetListImage({
        listKey: "company",
        itemId: id,
        fieldInternalName: fields.companyLogo,
        fileName: file.name,
        bytes,
        contentType: file.type || null,
      });
      const item = await getListItemByKey("company", id);
      if (!item) throw new ValidationError("Company not found after upload.");
      const company = mapCompanyFields(item.id, item.fields);
      return { company, logoUrl: company?.companyLogo ?? null };
    },
    { errorMessage: "Failed to upload company logo" },
    request,
  );
}
