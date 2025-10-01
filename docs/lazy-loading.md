# Lazy loading harness

The Lumina layout now exposes a shared entity loader that coordinates DOM skeletons
and asynchronous `google.script.run` calls. The harness is delivered by the new
`EntityLoader.html` partial and is included automatically for every view via
`layout.html`.

## Core concepts

- `LuminaEntityLoader.register(options)` registers a unit of work (an entity) and
  starts loading it once the DOM is ready. Every entity keeps a reference to the
  placeholder element, renders an optional skeleton, and runs your custom
  `load(context)` callback.
- `context.run(methodName, ...args)` is a promise wrapper around
  `google.script.run`. It rejects automatically when Apps Script is unavailable or
  when the requested server method is missing.
- Skeletons can be reused across pages by calling
  `LuminaEntityLoader.attachSkeleton(id, renderer)`. The renderer receives
  `{ id, placeholder, props, mount }` and should return HTML, a DOM node, or a
  fragment to display while the entity is loading.
- The loader ships with a default inline spinner skeleton registered as
  `lumina-inline-spinner`.

## Quick start

1. Make sure your page has a placeholder element inside the area that should show
   loading status and results. For example:

   ```html
   <div class="form-group">
     <select id="agentName"></select>
     <div id="agentLoaderHost" class="lumina-entity-placeholder" aria-live="polite"></div>
   </div>
   ```

2. Register the entity inside your DOM-ready handler:

   ```javascript
   LuminaEntityLoader.register({
     id: 'qualityForm.agentRoster',
     placeholder: document.getElementById('agentLoaderHost'),
     skeleton: { id: 'lumina-inline-spinner', props: { text: 'Loading agent roster…' } },
     load: async ({ run, mount, removeSkeleton }) => {
       try {
         const agents = await run('getUsers');
         mount(renderAgentOptions(agents));
       } finally {
         removeSkeleton();
       }
     }
   });
   ```

   The loader automatically clears the skeleton and mounts returned content. If
   your `load` callback returns a value (string, node, or fragment) the loader
   injects it into the placeholder. You can also call `mount()` manually and
   return `undefined` when you want full control.

3. When you need bespoke skeletons, register them once (anywhere after the
   partial loads):

   ```javascript
   LuminaEntityLoader.attachSkeleton('agents-pill', ({ props }) => {
     const wrapper = document.createElement('div');
     wrapper.className = 'pill-skeleton';
     wrapper.textContent = props.text || 'Fetching…';
     return wrapper;
   });
   ```

   Then point your entity at `skeleton: { id: 'agents-pill', props: { text: 'Loading…' } }`.

## Quality form example

`QualityForm.html` now registers `qualityForm.agentRoster` to hydrate the agent
selector. The loader handles:

- Showing the inline spinner placeholder
- Orchestrating the multi-step `google.script.run` fallbacks (`clientGetAssignedAgentNames`, `getUsers`, `getAllQA`)
- Surface-level messaging when the server is offline (the placeholder displays a
  warning/error status instead of an empty gap)

Review the script near the bottom of `QualityForm.html` for a full example of
mounting skeletons, calling `context.run`, and handling fallbacks.

## Additional helpers

The global loader also exposes:

- `LuminaEntityLoader.load(id)` – re-run a specific entity on demand
- `LuminaEntityLoader.loadAll()` – refresh every registered entity
- `LuminaEntityLoader.isGoogleScriptAvailable()` – feature detection for Apps
  Script availability
- `LuminaEntityLoader.run()` – direct access to the promise-wrapped
  `google.script.run` helper when you need it outside of an entity callback

These utilities make it easy to migrate existing feature pages to the shared
harness incrementally while reusing the same data access patterns.
