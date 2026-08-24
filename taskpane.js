/*
 * Original Image Size
 *
 * Shared-runtime version
 *
 * Automatic:
 *   Detect newly created images and restore them.
 *
 * Manual:
 *   Select an image and click "Restore Selected Image".
 */

let knownShapes = new Set();
let lastActivatedImage = null;
let monitorTimer = null;
let busy = false;
let initialized = false;


/* ============================================================
   OFFICE INITIALIZATION
   ============================================================ */

Office.onReady(async function () {

    console.log(
        "[Original Image Size] Office ready"
    );


    /*
     * Tell Excel to load this add-in automatically
     * whenever this workbook is opened.
     *
     * This does NOT open the task pane.
     */
    try {

        await Office.addin.setStartupBehavior(
            Office.StartupBehavior.load
        );

        console.log(
            "[Original Image Size] Startup behavior: LOAD"
        );

    }
    catch (error) {

        console.error(
            "[Original Image Size] Could not set startup behavior:",
            error
        );

    }


    /*
     * Task-pane button.
     */
    const button =
        document.getElementById("restore");

    if (button) {

        button.addEventListener(
            "click",
            restoreSelectedImage
        );

    }


    /*
     * Start the automatic image monitor.
     *
     * Because this is the shared runtime, this will
     * also run when Excel starts the add-in automatically.
     */
    await initializeMonitoring();

});


/* ============================================================
   RIBBON COMMAND REGISTRATION
   ============================================================ */

Office.actions.associate(
    "restoreSelectedImage",
    restoreSelectedImageCommand
);


/* ============================================================
   RIBBON: RESTORE SELECTED IMAGE
   ============================================================ */

async function restoreSelectedImageCommand(event) {

    try {

        await restoreSelectedImage();

    }
    catch (error) {

        console.error(
            "[Original Image Size] Ribbon command error:",
            error
        );

    }
    finally {

        /*
         * Required for ExecuteFunction commands.
         */
        event.completed();

    }
}


/* ============================================================
   INITIALIZE MONITORING
   ============================================================ */

async function initializeMonitoring() {

    /*
     * Already initialized.
     */
    if (initialized) {

        console.log(
            "[Original Image Size] Monitoring already initialized."
        );

        return;
    }


    /*
     * Extra protection:
     * never create a second timer.
     */
    if (monitorTimer !== null) {

        console.log(
            "[Original Image Size] Monitor timer already exists."
        );

        initialized = true;

        return;
    }


    initialized = true;


    try {

        await registerCurrentShapes();


        /*
         * Start exactly ONE monitor timer.
         */
        monitorTimer = setInterval(
            checkForNewShapes,
            200
        );


        setStatus(
            "✓ Automatic monitoring is running."
        );


        console.log(
            "[Original Image Size] Monitoring started."
        );

    }
    catch (error) {

        /*
         * If initialization failed before the timer
         * was created, allow a later retry.
         */
        initialized = false;


        console.error(
            "[Original Image Size] Initialization error:",
            error
        );


        setStatus(
            "Initialization error: " +
            (error.message || error)
        );
    }
}

/* ============================================================
   REGISTER EXISTING SHAPES
   ============================================================ */

async function registerCurrentShapes() {

    await Excel.run(async function (context) {

        const worksheets =
            context.workbook.worksheets;


        worksheets.load(
            "items/id"
        );


        await context.sync();


        knownShapes.clear();


        /*
         * Register shapes on every worksheet.
         */

        for (
            const worksheet
            of worksheets.items
        ) {

            const shapes =
                worksheet.shapes;


            shapes.load([
                "items/id",
                "items/type"
            ]);


            await context.sync();


            for (
                const shape
                of shapes.items
            ) {

                knownShapes.add(
                    makeShapeKey(
                        worksheet.id,
                        shape.id
                    )
                );


                /*
                 * Existing images should be available
                 * for manual restoration.
                 */

                if (
                    shape.type === "Image"
                ) {

                    registerActivationHandler(
                        shape
                    );

                }

            }

        }


        await context.sync();

    });

}


/* ============================================================
   REGISTER ACTIVATION HANDLER
   ============================================================ */

function registerActivationHandler(
    shape
) {

    try {

        shape.onActivated.add(
            rememberActivatedImage
        );

    }

    catch (error) {

        console.debug(
            "[Original Image Size] Could not register activation:",
            error
        );

    }

}


/* ============================================================
   REMEMBER ACTIVATED IMAGE
   ============================================================ */

async function rememberActivatedImage(
    event
) {

    try {

        /*
         * The event provides exactly the two IDs we need:
         * worksheetId + shapeId.
         */

        lastActivatedImage = {

            worksheetId:
                event.worksheetId,

            shapeId:
                event.shapeId

        };


        console.log(
            "[Original Image Size] Activated image:",
            event.shapeId
        );


        setStatus(
            "Image selected. Click Restore Selected Image."
        );

    }

    catch (error) {

        console.debug(
            "[Original Image Size] Activation error:",
            error
        );

    }

}


/* ============================================================
   CHECK FOR NEW SHAPES
   ============================================================ */

async function checkForNewShapes() {

    /*
     * Prevent overlapping Excel.run calls.
     */

    if (busy) {

        return;

    }


    busy = true;


    try {

        await Excel.run(async function (context) {

            const worksheets =
                context.workbook.worksheets;


            worksheets.load(
                "items/id"
            );


            await context.sync();


            for (
                const worksheet
                of worksheets.items
            ) {

                const shapes =
                    worksheet.shapes;


                shapes.load([
                    "items/id",
                    "items/type",
                    "items/name"
                ]);


                await context.sync();


                for (
                    const shape
                    of shapes.items
                ) {


                    const key =
                        makeShapeKey(
                            worksheet.id,
                            shape.id
                        );


                    /*
                     * We've already seen it.
                     */

                    if (
                        knownShapes.has(key)
                    ) {

                        continue;

                    }


                    /*
                     * Remember it immediately.
                     */

                    knownShapes.add(
                        key
                    );


                    console.log(
                        "[Original Image Size] New shape:",
                        shape.name,
                        shape.type
                    );


                    /*
                     * Ignore non-images.
                     */

                    if (
                        shape.type !== "Image"
                    ) {

                        continue;

                    }


                    /*
                     * Make the newly pasted image available
                     * to the manual button too.
                     */

                    registerActivationHandler(
                        shape
                    );


                    /*
                     * RESTORE ORIGINAL SIZE
                     */

                    setStatus(
                        "Restoring new image..."
                    );


                    shape.lockAspectRatio =
                        true;


                    shape.scaleWidth(
                        1,
                        Excel.ShapeScaleType.originalSize,
                        Excel.ShapeScaleFrom.topLeft
                    );


                    shape.scaleHeight(
                        1,
                        Excel.ShapeScaleType.originalSize,
                        Excel.ShapeScaleFrom.topLeft
                    );


                    await context.sync();


                    console.log(
                        "[Original Image Size] Restored:",
                        shape.name
                    );


                    setStatus(
                        "✓ New image restored."
                    );

                }

            }

        });

    }

    catch (error) {

        console.debug(
            "[Original Image Size] Monitor error:",
            error.message || error
        );

    }

    finally {

        busy = false;

    }

}


/* ============================================================
   MANUAL RESTORE
   ============================================================ */

async function restoreSelectedImage() {

    if (
        !lastActivatedImage
    ) {

        setStatus(
            "Please select an image first."
        );

        return;

    }


    try {

        setStatus(
            "Restoring selected image..."
        );


        await Excel.run(async function (context) {

            const worksheet =
                context.workbook.worksheets
                    .getItem(
                        lastActivatedImage
                            .worksheetId
                    );


            const shape =
                worksheet.shapes
                    .getItem(
                        lastActivatedImage
                            .shapeId
                    );


            shape.load([
                "id",
                "type",
                "name"
            ]);


            await context.sync();


            if (
                shape.type !== "Image"
            ) {

                throw new Error(
                    "The selected object is not an image."
                );

            }


            shape.lockAspectRatio =
                true;


            shape.scaleWidth(
                1,
                Excel.ShapeScaleType.originalSize,
                Excel.ShapeScaleFrom.topLeft
            );


            shape.scaleHeight(
                1,
                Excel.ShapeScaleType.originalSize,
                Excel.ShapeScaleFrom.topLeft
            );


            await context.sync();


            setStatus(
                "✓ Selected image restored."
            );


            console.log(
                "[Original Image Size] Manually restored:",
                shape.name
            );

        });

    }

    catch (error) {

        console.error(
            "[Original Image Size] Manual restore error:",
            error
        );


        setStatus(
            "Error: " +
            (error.message || error)
        );

    }

}


/* ============================================================
   SHAPE KEY
   ============================================================ */

function makeShapeKey(
    worksheetId,
    shapeId
) {

    return (
        worksheetId +
        "::" +
        shapeId
    );

}


/* ============================================================
   STATUS
   ============================================================ */

function setStatus(
    message
) {

    const element =
        document.getElementById(
            "status"
        );


    if (element) {

        element.textContent =
            message;

    }

}