SELECT oh.organization_id, oh.inventory_item_id, im.segment1 AS part_number, oh.qty
FROM ptq_OnHand_Sum AS oh LEFT JOIN ptq_ItemMap AS im ON (oh.organization_id = im.organization_id) AND (oh.inventory_item_id = im.inventory_item_id);

