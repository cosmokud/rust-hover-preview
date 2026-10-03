use std::path::Path;

use windows::core::{GUID, PCWSTR, VARIANT};
use windows::Win32::System::Com::{
    CLSIDFromProgID, CoCreateInstance, IDispatch, CLSCTX_LOCAL_SERVER, DISPATCH_FLAGS,
    DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO,
};

/// What a property put is identified by, in the parameter block that carries it.
const DISPID_PROPERTYPUT: i32 = -3;

// ----------------------------------------------------------- late binding

/// A late-bound automation object: an `IDispatch`, called by name.
///
/// Every Office application is available through `IDispatch`, so nothing here
/// depends on a type library, a build step or an interface this app would have to
/// keep in step with the installed Office. The price is that a parameter is named
/// as a string — and a name this Office does not know is dropped rather than
/// failing the call, since every parameter wanted here is optional and the
/// document's own name is the one that is not.
pub(super) struct Object(IDispatch);

impl Object {
    pub(super) fn create(prog_id: &str) -> Option<Self> {
        let wide: Vec<u16> = prog_id.encode_utf16().chain(std::iter::once(0)).collect();
        let class = unsafe { CLSIDFromProgID(PCWSTR(wide.as_ptr())).ok()? };
        let dispatch: IDispatch = unsafe {
            CoCreateInstance(
                &class,
                None::<&windows::core::IUnknown>,
                CLSCTX_LOCAL_SERVER,
            )
            .ok()?
        };

        Some(Self(dispatch))
    }

    pub(super) fn from_variant(value: VARIANT) -> Option<Self> {
        Some(Self(IDispatch::try_from(&value).ok()?))
    }

    fn dispatch_id(&self, name: &str) -> Option<i32> {
        let wide = wide_string(name);
        let mut id = 0i32;

        unsafe {
            self.0
                .GetIDsOfNames(&GUID::zeroed(), &PCWSTR(wide.as_ptr()), 1, 0, &mut id)
                .ok()?;
        }

        Some(id)
    }

    /// A parameter's dispatch ID, asked for with its member: the names go in as
    /// the member followed by the parameter, and the answer for the parameter is
    /// the second one.
    fn parameter_dispatch_id(&self, member: &str, parameter: &str) -> Option<i32> {
        let member = wide_string(member);
        let parameter = wide_string(parameter);
        let names = [PCWSTR(member.as_ptr()), PCWSTR(parameter.as_ptr())];
        let mut ids = [0i32; 2];

        unsafe {
            self.0
                .GetIDsOfNames(&GUID::zeroed(), names.as_ptr(), 2, 0, ids.as_mut_ptr())
                .ok()?;
        }

        Some(ids[1])
    }

    fn invoke(
        &self,
        name: &str,
        member: i32,
        flags: DISPATCH_FLAGS,
        values: &mut [VARIANT],
        ids: &mut [i32],
    ) -> Option<VARIANT> {
        let params = DISPPARAMS {
            rgvarg: if values.is_empty() {
                std::ptr::null_mut()
            } else {
                values.as_mut_ptr()
            },
            rgdispidNamedArgs: if ids.is_empty() {
                std::ptr::null_mut()
            } else {
                ids.as_mut_ptr()
            },
            cArgs: values.len() as u32,
            cNamedArgs: ids.len() as u32,
        };
        let mut result = VARIANT::new();
        // An automation failure carries its reason in the exception, not in the
        // HRESULT, so it is asked for rather than left behind.
        let mut exception = EXCEPINFO::default();

        let outcome = unsafe {
            self.0.Invoke(
                member,
                &GUID::zeroed(),
                0,
                flags,
                &params,
                Some(&mut result),
                Some(&mut exception),
                None,
            )
        };

        // A failed call is remembered with its reason: what it costs is one string
        // per failure, and what it buys is a document that can be looked at instead
        // of guessed about.
        if let Err(error) = &outcome {
            record_failure(name, error, &exception);
        }

        outcome.ok()?;
        Some(result)
    }

    /// A property's value.
    pub(super) fn value(&self, name: &str) -> Option<VARIANT> {
        let member = self.dispatch_id(name)?;
        self.invoke(name, member, DISPATCH_PROPERTYGET, &mut [], &mut [])
    }

    /// An item out of a collection, reached the way VBA reaches it when it writes
    /// `Slides(1)`.
    ///
    /// Some collections expose `Item` as a property and others as a method —
    /// PowerPoint's `Slides` is the second kind, and asking it as a property is
    /// answered with "member not found" — so the invoke says it may be either,
    /// which is what both kinds answer to.
    pub(super) fn item(&self, index: i32) -> Option<Self> {
        let member = self.dispatch_id("Item")?;
        let mut values = [VARIANT::from(index)];
        let value = self.invoke(
            "Item",
            member,
            DISPATCH_PROPERTYGET | DISPATCH_METHOD,
            &mut values,
            &mut [],
        )?;

        Self::from_variant(value)
    }

    /// A call with positional arguments, in the order they are written here: the
    /// parameter block carries them reversed, which is what the server expects.
    ///
    /// It is the way to reach the members whose parameter names do not resolve —
    /// `Range("A1:B2")` is one — and it needs no names to be right.
    pub(super) fn call_args(&self, name: &str, args: &[VARIANT]) -> Option<VARIANT> {
        let member = self.dispatch_id(name)?;
        let mut values: Vec<VARIANT> = args.iter().rev().cloned().collect();

        self.invoke(
            name,
            member,
            DISPATCH_PROPERTYGET | DISPATCH_METHOD,
            &mut values,
            &mut [],
        )
    }

    /// An object property.
    pub(super) fn member(&self, name: &str) -> Option<Self> {
        Self::from_variant(self.value(name)?)
    }

    pub(super) fn set(&self, name: &str, value: VARIANT) -> Option<()> {
        let member = self.dispatch_id(name)?;
        let mut values = [value];
        let mut ids = [DISPID_PROPERTYPUT];

        self.invoke(name, member, DISPATCH_PROPERTYPUT, &mut values, &mut ids)?;
        Some(())
    }

    /// A method call, with its arguments passed by name.
    pub(super) fn call(&self, name: &str, args: &[(&str, VARIANT)]) -> Option<VARIANT> {
        let member = self.dispatch_id(name)?;

        let mut values: Vec<VARIANT> = Vec::with_capacity(args.len());
        let mut ids: Vec<i32> = Vec::with_capacity(args.len());
        for (parameter, value) in args {
            let Some(id) = self.parameter_dispatch_id(name, parameter) else {
                continue;
            };
            ids.push(id);
            values.push(value.clone());
        }

        self.invoke(name, member, DISPATCH_METHOD, &mut values, &mut ids)
    }
}

fn wide_string(value: &str) -> Vec<u16> {
    value.encode_utf16().chain(std::iter::once(0)).collect()
}

pub(super) fn path_variant(path: &Path) -> VARIANT {
    VARIANT::from(path.to_string_lossy().as_ref())
}

thread_local! {
    /// What the last automation call failed with. It is kept because a render that
    /// produces nothing is otherwise silent — the file is simply left alone for a
    /// while — and this is what says which call refused and why: the diagnostic
    /// below prints it, and a failure remembered for a file carries it.
    static LAST_FAILURE: std::cell::RefCell<Option<String>> =
        const { std::cell::RefCell::new(None) };
}

fn record_failure(name: &str, error: &windows::core::Error, exception: &EXCEPINFO) {
    // An automation failure's reason is in the exception rather than in the
    // HRESULT, which for a failed call is only DISP_E_EXCEPTION.
    let description = exception.bstrDescription.to_string();
    let detail = if description.is_empty() {
        error.message()
    } else {
        format!("{} — {description}", error.message())
    };

    LAST_FAILURE.with(|slot| {
        *slot.borrow_mut() = Some(format!("{name}: 0x{:08X} {detail}", error.code().0));
    });
}

/// What the last automation call failed with, if one did. Read by the diagnostic
/// below: a render that produces nothing is otherwise silent.
#[cfg(test)]
pub(super) fn last_failure() -> Option<String> {
    LAST_FAILURE.with(|slot| slot.borrow().clone())
}
